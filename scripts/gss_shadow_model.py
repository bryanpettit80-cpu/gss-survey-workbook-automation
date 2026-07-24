#!/usr/bin/env python3
"""Privacy-safe, deterministic shadow modeling for GSS response exports.

The production command reads one JSON document from standard input and writes
only aggregate/model artifacts. Raw response rows, stable response identifiers,
comments, names, and contact data are never written to disk.

Input schema: ``gss-model-input/v1``

Required top-level fields:
  - source_sha256: SHA-256 of canonical JSON for the complete payload with the
    top-level source_sha256 field omitted
  - responses: response records defined by ALLOWED_RESPONSE_FIELDS
  - population_totals: authoritative current/previous 13-week aggregates

Optional top-level fields:
  - export_id: opaque export identifier (only its hash is persisted)
  - shadow_cycles_completed: advisory self-report retained for compatibility;
    it is never trusted for shadow readiness or promotion

Optional response field:
  - manager_visit: 1-5 rating accepted only when conditional_eligibility has
    manager_visit=true; used only in a separate sensitivity model

Each population total has:
  restaurant_id, window (current_13w or previous_13w), response_count,
  low_overall_rate_pct, and recommend_detractor_rate_pct.

This module intentionally has no file-input option. The only supported row-level
transport is stdin, and the only supported persistence is the five aggregate
artifacts documented in REQUIRED_ARTIFACTS.

Visit dates use local Monday-Sunday reporting weeks. Shadow readiness requires
eight distinct weeks in the external hash-chained aggregate cycle ledger; a
self-reported cycle count cannot produce ``ShadowReady``.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import importlib
import importlib.metadata
import io
import json
import math
import os
import re
import sys
import tempfile
import time
from contextlib import contextmanager, nullcontext
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Iterable, Iterator, Sequence

import numpy as np


INPUT_SCHEMA = "gss-model-input/v1"
SUMMARY_SCHEMA = "gss-shadow-model-summary/v1"
DIAGNOSTICS_SCHEMA = "gss-shadow-model-diagnostics/v1"
MANIFEST_SCHEMA = "gss-shadow-model-input-manifest/v1"
CYCLE_LEDGER_SCHEMA = "gss-shadow-cycle-ledger/v1"
EXPECTED_POLICY_VERSION = "gss-analysis-policy/v3"
TRACKED_POLICY_PATH = Path(__file__).resolve().parents[1] / "config" / "analysis-policy.json"

MANAGED_RUNTIME = {
    "numpy": ("numpy", "2.5.1"),
    "scipy": ("scipy", "1.18.0"),
    "scikit-learn": ("sklearn", "1.9.0"),
    "joblib": ("joblib", "1.5.3"),
    "threadpoolctl": ("threadpoolctl", "3.6.0"),
    "narwhals": ("narwhals", "2.24.0"),
}

REQUIRED_ARTIFACTS = (
    "model_summary.json",
    "model_estimates.csv",
    "model_diagnostics.json",
    "input_manifest.json",
    "model_card.md",
)

ALLOWED_TOP_LEVEL_FIELDS = {
    "schema_version",
    "source_sha256",
    "export_id",
    "shadow_cycles_completed",
    "responses",
    "population_totals",
}

ALLOWED_RESPONSE_FIELDS = {
    "response_id",
    "restaurant_id",
    "visit_date",
    "overall",
    "service",
    "culinary",
    "value",
    "pace",
    "recommend",
    "first_visit",
    "questionnaire_version",
    "conditional_eligibility",
    "manager_visit",
}

REQUIRED_RESPONSE_FIELDS = set(ALLOWED_RESPONSE_FIELDS) - {"manager_visit"}

ALLOWED_POPULATION_FIELDS = {
    "restaurant_id",
    "window",
    "response_count",
    "low_overall_rate_pct",
    "recommend_detractor_rate_pct",
}

DRIVERS = ("service", "culinary", "value", "pace")
DRIVER_LABELS = {
    "service": "Service",
    "culinary": "Culinary",
    "value": "Value",
    "pace": "Pace",
}

OUTCOMES = {
    "low_overall": {
        "policy_name": "overall_rating_1_to_3",
        "csv_name": "low_overall_1_to_3",
        "label": "Low Overall (1-3)",
    },
    "recommend_detractor": {
        "policy_name": "recommend_rating_0_to_6",
        "csv_name": "recommend_detractor_0_to_6",
        "label": "Recommend detractor (0-6)",
    },
}

WINDOWS = ("current_13w", "previous_13w")


class DataBlock(Exception):
    """A safe-to-report condition that prevents inferential analysis."""

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message


@dataclass(frozen=True)
class Response:
    response_id: str
    restaurant_id: str
    visit_date: date
    week_start: date
    overall: float | None
    service: float | None
    culinary: float | None
    value: float | None
    pace: float | None
    recommend: float | None
    first_visit: bool
    questionnaire_version: str
    manager_visit: float | None

    @property
    def low_overall(self) -> int | None:
        if self.overall is None:
            return None
        return int(self.overall <= 3)

    @property
    def recommend_detractor(self) -> int | None:
        if self.recommend is None:
            return None
        return int(self.recommend <= 6)

    @property
    def model_usable(self) -> bool:
        return self.model_usable_for("low_overall")

    def outcome_value(self, outcome: str) -> int | None:
        if outcome == "low_overall":
            return self.low_overall
        if outcome == "recommend_detractor":
            return self.recommend_detractor
        raise ValueError(f"Unsupported outcome: {outcome}")

    def model_usable_for(
        self,
        outcome: str,
        include_manager_visit: bool = False,
    ) -> bool:
        return (
            self.outcome_value(outcome) is not None
            and self.service is not None
            and self.culinary is not None
            and self.value is not None
            and self.pace is not None
            and (not include_manager_visit or self.manager_visit is not None)
        )


@dataclass(frozen=True)
class PopulationTotal:
    restaurant_id: str
    window: str
    response_count: int
    low_overall_rate_pct: float
    recommend_detractor_rate_pct: float


@dataclass(frozen=True)
class DesignSpec:
    restaurant_levels: tuple[str, ...]
    questionnaire_levels: tuple[str, ...]
    first_week: date
    feature_names: tuple[str, ...]


def canonical_json(value: Any) -> str:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=True)


def canonical_source_sha256(payload: Any) -> str:
    """Hash the complete model payload without the recursive hash field."""

    if not isinstance(payload, dict):
        raise DataBlock(
            "invalid_input",
            "The stdin JSON document must be an object.",
        )
    hashable_payload = {
        key: value for key, value in payload.items() if key != "source_sha256"
    }
    return sha256_text(canonical_json(hashable_payload))


def managed_runtime_versions() -> dict[str, str | None]:
    versions: dict[str, str | None] = {}
    for distribution_name in MANAGED_RUNTIME:
        try:
            versions[distribution_name] = importlib.metadata.version(
                distribution_name
            )
        except importlib.metadata.PackageNotFoundError:
            versions[distribution_name] = None
    return versions


def ensure_managed_runtime() -> None:
    actual_versions = managed_runtime_versions()
    failures = []
    for distribution_name, (module_name, required_version) in MANAGED_RUNTIME.items():
        actual_version = actual_versions[distribution_name]
        if actual_version != required_version:
            failures.append(
                f"{distribution_name}=={required_version} required; "
                f"found {actual_version or 'not installed'}"
            )
            continue
        try:
            importlib.import_module(module_name)
        except ImportError:
            failures.append(
                f"{distribution_name}=={required_version} is installed but cannot be imported"
            )
    if failures:
        raise RuntimeError(
            "Managed GSS modeling runtime is incomplete: " + "; ".join(failures)
        )


def sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def round_or_none(value: float | None, digits: int = 8) -> float | None:
    if value is None or not math.isfinite(value):
        return None
    return round(float(value), digits)


def monday_week_start(value: date) -> date:
    """Return the Monday for the GSS Monday-Sunday reporting week."""

    return value - timedelta(days=value.weekday())


def parse_iso_date(value: Any, field: str) -> date:
    if not isinstance(value, str) or not value.strip():
        raise DataBlock("invalid_date", f"{field} must be a non-empty ISO local date.")
    try:
        parsed = date.fromisoformat(value)
    except ValueError as exc:
        raise DataBlock("invalid_date", f"{field} must use YYYY-MM-DD.") from exc
    return parsed


def parse_optional_number(
    value: Any,
    field: str,
    minimum: float,
    maximum: float,
) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise DataBlock("invalid_score", f"{field} must be numeric or null.")
    numeric = float(value)
    if not math.isfinite(numeric) or numeric < minimum or numeric > maximum:
        raise DataBlock(
            "invalid_score",
            f"{field} must be between {minimum:g} and {maximum:g}, or null.",
        )
    return numeric


def parse_bool(value: Any, field: str) -> bool:
    if isinstance(value, bool):
        return value
    if value in (0, 1):
        return bool(value)
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"yes", "y", "true", "1"}:
            return True
        if normalized in {"no", "n", "false", "0"}:
            return False
    raise DataBlock("invalid_boolean", f"{field} must be a boolean.")


def validate_identifier(value: Any, field: str, maximum_length: int = 128) -> str:
    if not isinstance(value, str) or not value.strip():
        raise DataBlock("missing_identifier", f"{field} must be non-empty.")
    normalized = value.strip()
    if len(normalized) > maximum_length:
        raise DataBlock("invalid_identifier", f"{field} exceeds the length limit.")
    return normalized


def validate_conditional_eligibility(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if not isinstance(value, dict):
        raise DataBlock(
            "invalid_conditional_eligibility",
            "conditional_eligibility must be a boolean or a flag object.",
        )
    if len(value) > 64:
        raise DataBlock(
            "invalid_conditional_eligibility",
            "conditional_eligibility contains too many flags.",
        )
    for key, flag in value.items():
        if (
            not isinstance(key, str)
            or not re.fullmatch(r"[A-Za-z0-9_.-]{1,64}", key)
            or not isinstance(flag, bool)
        ):
            raise DataBlock(
                "invalid_conditional_eligibility",
                "conditional_eligibility objects may contain only named boolean flags.",
            )
    return bool(value.get("manager_visit", False))


def parse_response(raw: Any) -> Response:
    if not isinstance(raw, dict):
        raise DataBlock("invalid_response", "Every response must be a JSON object.")

    unexpected = sorted(set(raw) - ALLOWED_RESPONSE_FIELDS)
    if unexpected:
        raise DataBlock(
            "privacy_unexpected_fields",
            "Response rows contain one or more fields outside the privacy allowlist.",
        )
    missing = sorted(REQUIRED_RESPONSE_FIELDS - set(raw))
    if missing:
        raise DataBlock(
            "missing_required_fields",
            "Response rows are missing required fields: " + ", ".join(missing),
        )

    manager_visit_eligible = validate_conditional_eligibility(
        raw["conditional_eligibility"]
    )
    manager_visit = parse_optional_number(
        raw.get("manager_visit"),
        "manager_visit",
        1,
        5,
    )
    if manager_visit is not None and not manager_visit_eligible:
        raise DataBlock(
            "manager_visit_not_eligible",
            "manager_visit may be populated only when its conditional eligibility flag is true.",
        )

    visit_date = parse_iso_date(raw["visit_date"], "visit_date")
    questionnaire_version = validate_identifier(
        raw["questionnaire_version"],
        "questionnaire_version",
        64,
    )

    return Response(
        response_id=validate_identifier(raw["response_id"], "response_id"),
        restaurant_id=validate_identifier(raw["restaurant_id"], "restaurant_id", 64),
        visit_date=visit_date,
        week_start=monday_week_start(visit_date),
        overall=parse_optional_number(raw["overall"], "overall", 1, 5),
        service=parse_optional_number(raw["service"], "service", 1, 5),
        culinary=parse_optional_number(raw["culinary"], "culinary", 1, 5),
        value=parse_optional_number(raw["value"], "value", 1, 5),
        pace=parse_optional_number(raw["pace"], "pace", 1, 5),
        recommend=parse_optional_number(raw["recommend"], "recommend", 0, 10),
        first_visit=parse_bool(raw["first_visit"], "first_visit"),
        questionnaire_version=questionnaire_version,
        manager_visit=manager_visit,
    )


def parse_population_total(raw: Any) -> PopulationTotal:
    if not isinstance(raw, dict):
        raise DataBlock(
            "invalid_population_total",
            "Every population total must be a JSON object.",
        )
    unexpected = sorted(set(raw) - ALLOWED_POPULATION_FIELDS)
    missing = sorted(ALLOWED_POPULATION_FIELDS - set(raw))
    if unexpected:
        raise DataBlock(
            "unexpected_population_fields",
            "Population totals contain one or more unsupported fields.",
        )
    if missing:
        raise DataBlock(
            "missing_population_fields",
            "Population totals are missing fields: " + ", ".join(missing),
        )

    window = raw["window"]
    if window not in WINDOWS:
        raise DataBlock(
            "invalid_population_window",
            "population_totals.window must be current_13w or previous_13w.",
        )
    count = raw["response_count"]
    if isinstance(count, bool) or not isinstance(count, int) or count < 0:
        raise DataBlock(
            "invalid_population_count",
            "population_totals.response_count must be a non-negative integer.",
        )

    def rate(name: str) -> float:
        value = raw[name]
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise DataBlock(
                "invalid_population_rate",
                f"population_totals.{name} must be numeric.",
            )
        numeric = float(value)
        if not math.isfinite(numeric) or numeric < 0 or numeric > 100:
            raise DataBlock(
                "invalid_population_rate",
                f"population_totals.{name} must be between 0 and 100.",
            )
        return numeric

    return PopulationTotal(
        restaurant_id=validate_identifier(
            raw["restaurant_id"],
            "population_totals.restaurant_id",
            64,
        ),
        window=window,
        response_count=count,
        low_overall_rate_pct=rate("low_overall_rate_pct"),
        recommend_detractor_rate_pct=rate("recommend_detractor_rate_pct"),
    )


def parse_payload(payload: Any) -> tuple[list[Response], list[PopulationTotal]]:
    if not isinstance(payload, dict):
        raise DataBlock("invalid_input", "The stdin JSON document must be an object.")

    unexpected = sorted(set(payload) - ALLOWED_TOP_LEVEL_FIELDS)
    if unexpected:
        raise DataBlock(
            "unexpected_top_level_fields",
            "Input contains one or more unsupported top-level fields.",
        )
    if payload.get("schema_version") != INPUT_SCHEMA:
        raise DataBlock(
            "invalid_schema_version",
            f"schema_version must be {INPUT_SCHEMA}.",
        )

    source_hash = payload.get("source_sha256")
    if (
        not isinstance(source_hash, str)
        or len(source_hash) != 64
        or any(character not in "0123456789abcdefABCDEF" for character in source_hash)
    ):
        raise DataBlock(
            "invalid_source_hash",
            "source_sha256 must be a 64-character hexadecimal SHA-256.",
        )
    expected_source_hash = canonical_source_sha256(payload)
    if source_hash.lower() != expected_source_hash:
        raise DataBlock(
            "source_hash_mismatch",
            "source_sha256 does not match the canonical payload with source_sha256 omitted.",
        )

    raw_responses = payload.get("responses")
    if not isinstance(raw_responses, list) or not raw_responses:
        raise DataBlock(
            "missing_responses",
            "responses must be a non-empty JSON array.",
        )
    raw_totals = payload.get("population_totals")
    if not isinstance(raw_totals, list) or not raw_totals:
        raise DataBlock(
            "missing_population_totals",
            "population_totals must be a non-empty JSON array.",
        )

    responses = [parse_response(raw) for raw in raw_responses]
    totals = [parse_population_total(raw) for raw in raw_totals]
    return responses, totals


def response_record_set_hash(
    responses: Sequence[Response],
    totals: Sequence[PopulationTotal],
) -> str:
    safe_records = []
    for response in sorted(responses, key=lambda item: item.response_id):
        safe_records.append(
            {
                "response_id_sha256": sha256_text(response.response_id),
                "restaurant_id": response.restaurant_id,
                "visit_date": response.visit_date.isoformat(),
                "overall": response.overall,
                "service": response.service,
                "culinary": response.culinary,
                "value": response.value,
                "pace": response.pace,
                "recommend": response.recommend,
                "first_visit": response.first_visit,
                "questionnaire_version": response.questionnaire_version,
                "manager_visit": response.manager_visit,
            }
        )
    safe_totals = [
        {
            "restaurant_id": total.restaurant_id,
            "window": total.window,
            "response_count": total.response_count,
            "low_overall_rate_pct": total.low_overall_rate_pct,
            "recommend_detractor_rate_pct": total.recommend_detractor_rate_pct,
        }
        for total in sorted(
            totals,
            key=lambda item: (item.restaurant_id, item.window),
        )
    ]
    return sha256_text(
        canonical_json(
            {
                "responses": safe_records,
                "population_totals": safe_totals,
            }
        )
    )


def load_tracked_policy() -> dict[str, Any]:
    try:
        raw_bytes = TRACKED_POLICY_PATH.read_bytes()
        tracked = json.loads(raw_bytes.decode("utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise RuntimeError(
            "The tracked GSS analysis policy could not be loaded."
        ) from exc
    if tracked.get("schema_version") != EXPECTED_POLICY_VERSION:
        raise RuntimeError(
            f"The tracked GSS analysis policy must be {EXPECTED_POLICY_VERSION}."
        )
    try:
        population = tracked["population_export"]
        reconciliation = population["reconciliation"]
        inference = tracked["inference"]
        comparison = inference["comparison"]
        driver = tracked["driver_model"]
        validation = driver["validation"]
        shadow = driver["shadow_minimums"]
        store = driver["store_specific_minimums"]
        promotion = driver["promotion"]
        forecasting = tracked["forecasting"]
    except (KeyError, TypeError) as exc:
        raise RuntimeError(
            "The tracked GSS analysis policy is missing required modeling keys."
        ) from exc

    required_fields = set(population.get("required_fields", []))
    optional_fields = set(population.get("optional_fields", []))
    if required_fields != REQUIRED_RESPONSE_FIELDS or optional_fields != {
        "manager_visit"
    }:
        raise RuntimeError(
            "The tracked population-export fields do not match the model input contract."
        )
    if comparison.get("interval_method") != "newcombe_independent_score":
        raise RuntimeError(
            "The tracked inference interval method must be newcombe_independent_score."
        )
    if population.get("source_sha256_definition") != (
        "sha256_canonical_json_omitting_source_sha256"
    ):
        raise RuntimeError(
            "The tracked source SHA-256 definition does not match the model contract."
        )
    if population.get("reporting_week") != {
        "start_day": "Monday",
        "end_day": "Sunday",
        "date_basis": "local_visit_date",
    }:
        raise RuntimeError(
            "The tracked population export must use local Monday-Sunday weeks."
        )
    if not driver.get("manager_visit_sensitivity_only"):
        raise RuntimeError(
            "The tracked policy must keep Manager Visit sensitivity-only."
        )
    if (
        inference.get("primary_outcome")
        != {
            "field": "overall",
            "event_values": [1, 2, 3],
            "label": "low_overall",
        }
        or inference.get("secondary_outcome")
        != {
            "field": "recommend",
            "event_minimum": 0,
            "event_maximum": 6,
            "label": "recommend_detractor",
        }
        or driver.get("primary_outcome") != "low_overall"
        or driver.get("secondary_outcome") != "recommend_detractor"
    ):
        raise RuntimeError(
            "The tracked policy outcomes do not match the implemented model contract."
        )
    if comparison.get("windows_overlap") is not False or comparison.get(
        "current_window_weeks"
    ) != comparison.get("previous_window_weeks"):
        raise RuntimeError(
            "The tracked policy must use equal non-overlapping comparison windows."
        )
    if promotion.get("calibration_model_must_converge") is not True:
        raise RuntimeError(
            "The tracked promotion policy must require calibration convergence."
        )
    if shadow.get("cycle_evidence") != (
        "verified_append_only_aggregate_ledger"
    ):
        raise RuntimeError(
            "The tracked shadow gate must require durable aggregate cycle evidence."
        )

    return {
        "version": tracked["schema_version"],
        "source_sha256": hashlib.sha256(raw_bytes).hexdigest(),
        "source_path": "config/analysis-policy.json",
        "primary_outcome": driver["primary_outcome"],
        "secondary_outcome": driver["secondary_outcome"],
        "reconciliation": {
            "count_relative_tolerance": reconciliation[
                "count_delta_maximum_rate"
            ],
            "count_absolute_tolerance": reconciliation[
                "count_delta_minimum_absolute"
            ],
            "score_tolerance_percentage_points": reconciliation[
                "score_delta_maximum_percentage_points"
            ],
            "minimum_core_completeness": reconciliation[
                "core_rating_completeness_minimum_rate"
            ],
        },
        "alerts": {
            "window_weeks": comparison["current_window_weeks"],
            "minimum_responses_each_window": comparison[
                "minimum_responses_per_window"
            ],
            "minimum_practical_change_percentage_points": comparison[
                "minimum_practical_difference_percentage_points"
            ],
            "false_discovery_rate": comparison["false_discovery_rate"],
            "confidence_level": comparison["confidence_level"],
        },
        "model": {
            "c_grid": driver["regularization_c_grid"],
            "validation_block_weeks": validation[
                "expanding_validation_block_weeks"
            ],
            "holdout_weeks": validation["holdout_weeks"],
            "minimum_training_weeks_before_fold": validation[
                "minimum_training_weeks_before_fold"
            ],
            "bootstrap_replicates": validation["bootstrap_repetitions"],
            "bootstrap_seed": validation["random_seed"],
            "all_folds_must_converge": validation[
                "all_folds_must_converge"
            ],
            "final_model_must_converge": validation[
                "final_model_must_converge"
            ],
            "all_bootstrap_repetitions_required": validation[
                "all_bootstrap_repetitions_required"
            ],
        },
        "shadow_gate": {
            "minimum_usable_responses": shadow["usable_responses"],
            "minimum_primary_events": shadow["primary_events"],
            "minimum_distinct_weeks": shadow["weeks"],
            "minimum_cycles": shadow["distinct_weekly_cycles"],
            "store_minimum_usable_responses": store[
                "usable_responses_per_store"
            ],
            "store_minimum_primary_events": store[
                "primary_events_per_store"
            ],
        },
        "promotion": {
            "minimum_distinct_weeks": promotion["minimum_weeks"],
            "minimum_usable_responses": promotion["minimum_responses"],
            "minimum_primary_events": promotion["minimum_primary_events"],
            "minimum_validation_folds": promotion[
                "minimum_validation_folds"
            ],
            "minimum_brier_improvement_fraction": promotion[
                "minimum_brier_improvement_rate"
            ],
            "minimum_auc": promotion["minimum_auc"],
            "calibration_slope_minimum": promotion[
                "calibration_slope_minimum"
            ],
            "calibration_slope_maximum": promotion[
                "calibration_slope_maximum"
            ],
            "maximum_absolute_calibration_intercept": promotion[
                "calibration_intercept_absolute_maximum"
            ],
            "minimum_top_factor_sign_stability": promotion[
                "top_factor_sign_stability_minimum_rate"
            ],
            "minimum_shadow_cycles": shadow["distinct_weekly_cycles"],
        },
        "forecasting": {
            "minimum_contiguous_weeks": forecasting[
                "minimum_contiguous_population_weeks"
            ],
            "minimum_seasonality_weeks": forecasting[
                "minimum_weeks_for_annual_seasonality"
            ],
        },
    }


def policy_with_override(policy_override: dict[str, Any] | None) -> dict[str, Any]:
    policy = load_tracked_policy()
    if not policy_override:
        return policy
    allowed = {"bootstrap_replicates"}
    unexpected = set(policy_override) - allowed
    if unexpected:
        raise ValueError(f"Unsupported internal policy override: {sorted(unexpected)}")
    if "bootstrap_replicates" in policy_override:
        value = policy_override["bootstrap_replicates"]
        if not isinstance(value, int) or value < 1:
            raise ValueError("bootstrap_replicates must be a positive integer")
        policy["model"]["bootstrap_replicates"] = value
    return policy


def get_window_bounds(max_week: date) -> dict[str, tuple[date, date]]:
    current_start = max_week - timedelta(weeks=12)
    previous_end = current_start - timedelta(weeks=1)
    previous_start = previous_end - timedelta(weeks=12)
    return {
        "current_13w": (current_start, max_week),
        "previous_13w": (previous_start, previous_end),
    }


def responses_in_window(
    responses: Sequence[Response],
    restaurant_id: str,
    bounds: tuple[date, date],
) -> list[Response]:
    start, end = bounds
    return [
        response
        for response in responses
        if response.restaurant_id == restaurant_id
        and start <= response.week_start <= end
    ]


def rate_percentage(values: Iterable[int | None]) -> tuple[int, float | None]:
    usable = [value for value in values if value is not None]
    if not usable:
        return 0, None
    return len(usable), (sum(usable) / len(usable)) * 100.0


def reconcile_population(
    responses: Sequence[Response],
    totals: Sequence[PopulationTotal],
    policy: dict[str, Any],
) -> dict[str, Any]:
    duplicate_ids = len(responses) - len({response.response_id for response in responses})
    restaurants = sorted({response.restaurant_id for response in responses})
    max_week = max(response.week_start for response in responses)
    bounds = get_window_bounds(max_week)
    reconciliation_policy = policy["reconciliation"]

    target_by_key: dict[tuple[str, str], PopulationTotal] = {}
    duplicate_targets: list[str] = []
    for target in totals:
        key = (target.restaurant_id, target.window)
        if key in target_by_key:
            duplicate_targets.append(f"{target.restaurant_id}/{target.window}")
        target_by_key[key] = target

    failures: list[dict[str, Any]] = []
    if duplicate_ids:
        failures.append(
            {
                "code": "duplicate_response_ids",
                "count": duplicate_ids,
            }
        )
    if duplicate_targets:
        failures.append(
            {
                "code": "duplicate_population_targets",
                "count": len(duplicate_targets),
            }
        )

    details: list[dict[str, Any]] = []
    core_present = 0
    core_expected = len(responses) * 6
    for response in responses:
        core_present += sum(
            value is not None
            for value in (
                response.overall,
                response.service,
                response.culinary,
                response.value,
                response.pace,
                response.recommend,
            )
        )
    core_completeness = core_present / core_expected if core_expected else 0.0
    if core_completeness < reconciliation_policy["minimum_core_completeness"]:
        failures.append(
            {
                "code": "core_completeness_below_minimum",
                "actual": round_or_none(core_completeness, 6),
                "minimum": reconciliation_policy["minimum_core_completeness"],
            }
        )

    expected_keys = {
        (restaurant_id, window)
        for restaurant_id in restaurants
        for window in WINDOWS
    }
    extra_keys = set(target_by_key) - expected_keys
    if extra_keys:
        failures.append(
            {
                "code": "population_targets_outside_export",
                "count": len(extra_keys),
            }
        )

    for restaurant_id in restaurants:
        for window in WINDOWS:
            rows = responses_in_window(
                responses,
                restaurant_id,
                bounds[window],
            )
            target = target_by_key.get((restaurant_id, window))
            if target is None:
                failures.append(
                    {
                        "code": "missing_population_target",
                        "restaurant_id": restaurant_id,
                        "window": window,
                    }
                )
                continue

            overall_n, overall_rate = rate_percentage(
                response.low_overall for response in rows
            )
            recommend_n, recommend_rate = rate_percentage(
                response.recommend_detractor for response in rows
            )
            tolerance = max(
                reconciliation_policy["count_absolute_tolerance"],
                target.response_count
                * reconciliation_policy["count_relative_tolerance"],
            )
            count_delta = len(rows) - target.response_count
            window_core_expected = len(rows) * 6
            window_core_present = sum(
                value is not None
                for response in rows
                for value in (
                    response.overall,
                    response.service,
                    response.culinary,
                    response.value,
                    response.pace,
                    response.recommend,
                )
            )
            window_core_completeness = (
                window_core_present / window_core_expected
                if window_core_expected
                else 1.0
            )
            overall_delta = (
                None
                if overall_rate is None
                else overall_rate - target.low_overall_rate_pct
            )
            recommend_delta = (
                None
                if recommend_rate is None
                else recommend_rate - target.recommend_detractor_rate_pct
            )
            passed = (
                abs(count_delta) <= tolerance
                and overall_delta is not None
                and abs(overall_delta)
                <= reconciliation_policy["score_tolerance_percentage_points"]
                and recommend_delta is not None
                and abs(recommend_delta)
                <= reconciliation_policy["score_tolerance_percentage_points"]
                and window_core_completeness
                >= reconciliation_policy["minimum_core_completeness"]
            )
            detail = {
                "restaurant_id": restaurant_id,
                "window": window,
                "export_response_count": len(rows),
                "population_response_count": target.response_count,
                "count_delta": count_delta,
                "count_tolerance": round_or_none(tolerance, 4),
                "core_completeness": round_or_none(
                    window_core_completeness,
                    6,
                ),
                "overall_scored_count": overall_n,
                "low_overall_rate_pct": round_or_none(overall_rate, 6),
                "low_overall_rate_delta_pp": round_or_none(overall_delta, 6),
                "recommend_scored_count": recommend_n,
                "recommend_detractor_rate_pct": round_or_none(
                    recommend_rate,
                    6,
                ),
                "recommend_detractor_rate_delta_pp": round_or_none(
                    recommend_delta,
                    6,
                ),
                "passed": passed,
            }
            details.append(detail)
            if not passed:
                failures.append(
                    {
                        "code": "population_reconciliation_mismatch",
                        "restaurant_id": restaurant_id,
                        "window": window,
                    }
                )

    return {
        "status": "Passed" if not failures else "DataBlocked",
        "max_response_week": max_week.isoformat(),
        "windows": {
            window: {
                "start_week": start.isoformat(),
                "end_week": end.isoformat(),
            }
            for window, (start, end) in bounds.items()
        },
        "response_id_unique": duplicate_ids == 0,
        "duplicate_response_id_count": duplicate_ids,
        "core_completeness": round_or_none(core_completeness, 6),
        "date_completeness": 1.0,
        "restaurant_completeness": 1.0,
        "details": details,
        "failures": failures,
    }


def wilson_interval(successes: int, total: int, z: float = 1.959963984540054) -> tuple[float, float]:
    if total <= 0:
        return math.nan, math.nan
    proportion = successes / total
    z2 = z * z
    denominator = 1 + z2 / total
    center = (proportion + z2 / (2 * total)) / denominator
    half_width = (
        z
        * math.sqrt(
            proportion * (1 - proportion) / total + z2 / (4 * total * total)
        )
        / denominator
    )
    return max(0.0, center - half_width), min(1.0, center + half_width)


def newcombe_independent_difference_interval(
    current_successes: int,
    current_total: int,
    previous_successes: int,
    previous_total: int,
    z: float = 1.959963984540054,
) -> tuple[float, float]:
    """Newcombe score interval for two independent binomial proportions.

    This is Newcombe's hybrid score method without continuity correction. The
    lower and upper limits combine the appropriate Wilson score distances in
    quadrature; subtracting the two Wilson endpoints is not the Newcombe
    interval and is unnecessarily conservative.
    """

    if current_total <= 0 or previous_total <= 0:
        return math.nan, math.nan
    current_rate = current_successes / current_total
    previous_rate = previous_successes / previous_total
    difference = current_rate - previous_rate
    current_low, current_high = wilson_interval(
        current_successes,
        current_total,
        z,
    )
    previous_low, previous_high = wilson_interval(
        previous_successes,
        previous_total,
        z,
    )
    lower = difference - math.sqrt(
        (current_rate - current_low) ** 2
        + (previous_high - previous_rate) ** 2
    )
    upper = difference + math.sqrt(
        (current_high - current_rate) ** 2
        + (previous_rate - previous_low) ** 2
    )
    return max(-1.0, lower), min(1.0, upper)


def two_proportion_p_value(
    current_successes: int,
    current_total: int,
    previous_successes: int,
    previous_total: int,
) -> float:
    if current_total <= 0 or previous_total <= 0:
        return 1.0
    pooled = (current_successes + previous_successes) / (
        current_total + previous_total
    )
    standard_error = math.sqrt(
        pooled
        * (1 - pooled)
        * (1 / current_total + 1 / previous_total)
    )
    if standard_error == 0:
        return 1.0
    difference = (
        current_successes / current_total
        - previous_successes / previous_total
    )
    z_score = abs(difference) / standard_error
    return math.erfc(z_score / math.sqrt(2))


def add_bh_q_values(items: list[dict[str, Any]]) -> None:
    if not items:
        return
    ordered = sorted(enumerate(items), key=lambda pair: pair[1]["p_value"])
    count = len(ordered)
    adjusted = [1.0] * count
    running = 1.0
    for reverse_index in range(count - 1, -1, -1):
        original_index, item = ordered[reverse_index]
        rank = reverse_index + 1
        candidate = min(1.0, item["p_value"] * count / rank)
        running = min(running, candidate)
        adjusted[original_index] = running
    for index, item in enumerate(items):
        item["q_value"] = round_or_none(adjusted[index], 10)


def build_alerts(
    responses: Sequence[Response],
    reconciliation: dict[str, Any],
    policy: dict[str, Any],
) -> list[dict[str, Any]]:
    bounds = {
        window: (
            date.fromisoformat(values["start_week"]),
            date.fromisoformat(values["end_week"]),
        )
        for window, values in reconciliation["windows"].items()
    }
    alert_policy = policy["alerts"]
    items: list[dict[str, Any]] = []
    restaurants = sorted({response.restaurant_id for response in responses})
    outcome_accessors = (
        ("low_overall", lambda item: item.low_overall),
        ("recommend_detractor", lambda item: item.recommend_detractor),
    )

    for restaurant_id in restaurants:
        current_rows = responses_in_window(
            responses,
            restaurant_id,
            bounds["current_13w"],
        )
        previous_rows = responses_in_window(
            responses,
            restaurant_id,
            bounds["previous_13w"],
        )
        for outcome, accessor in outcome_accessors:
            current_values = [
                value
                for value in (accessor(response) for response in current_rows)
                if value is not None
            ]
            previous_values = [
                value
                for value in (accessor(response) for response in previous_rows)
                if value is not None
            ]
            current_total = len(current_values)
            previous_total = len(previous_values)
            current_successes = sum(current_values)
            previous_successes = sum(previous_values)
            current_rate = (
                current_successes / current_total if current_total else math.nan
            )
            previous_rate = (
                previous_successes / previous_total if previous_total else math.nan
            )
            difference = current_rate - previous_rate
            difference_ci = newcombe_independent_difference_interval(
                current_successes,
                current_total,
                previous_successes,
                previous_total,
            )
            items.append(
                {
                    "restaurant_id": restaurant_id,
                    "outcome": outcome,
                    "current_n": current_total,
                    "previous_n": previous_total,
                    "current_rate_pct": round_or_none(current_rate * 100, 6),
                    "previous_rate_pct": round_or_none(previous_rate * 100, 6),
                    "change_percentage_points": round_or_none(
                        difference * 100,
                        6,
                    ),
                    "newcombe_95_ci_percentage_points": [
                        round_or_none(difference_ci[0] * 100, 6),
                        round_or_none(difference_ci[1] * 100, 6),
                    ],
                    "p_value": round_or_none(
                        two_proportion_p_value(
                            current_successes,
                            current_total,
                            previous_successes,
                            previous_total,
                        ),
                        10,
                    ),
                }
            )

    add_bh_q_values(items)
    for item in items:
        sample_eligible = (
            item["current_n"]
            >= alert_policy["minimum_responses_each_window"]
            and item["previous_n"]
            >= alert_policy["minimum_responses_each_window"]
        )
        practical = (
            item["change_percentage_points"] is not None
            and abs(item["change_percentage_points"])
            >= alert_policy["minimum_practical_change_percentage_points"]
        )
        statistically_supported = (
            item["q_value"] is not None
            and item["q_value"] <= alert_policy["false_discovery_rate"]
        )
        item["sample_eligible"] = sample_eligible
        item["practical_threshold_met"] = practical
        item["fdr_threshold_met"] = statistically_supported
        item["alert"] = sample_eligible and practical and statistically_supported
        if item["change_percentage_points"] is None:
            item["direction"] = "not_scored"
        elif item["change_percentage_points"] > 0:
            item["direction"] = "worsening"
        elif item["change_percentage_points"] < 0:
            item["direction"] = "improving"
        else:
            item["direction"] = "unchanged"
    return items


def longest_contiguous_week_run(weeks: Sequence[date]) -> int:
    if not weeks:
        return 0
    ordered = sorted(set(weeks))
    longest = 1
    current = 1
    for previous, value in zip(ordered, ordered[1:]):
        if value - previous == timedelta(weeks=1):
            current += 1
        else:
            current = 1
        longest = max(longest, current)
    return longest


def make_design_spec(
    responses: Sequence[Response],
    include_manager_visit: bool = False,
) -> DesignSpec:
    restaurant_levels = tuple(sorted({item.restaurant_id for item in responses}))
    questionnaire_levels = tuple(
        sorted({item.questionnaire_version for item in responses})
    )
    feature_names = ["intercept", *(f"{driver}_worsening" for driver in DRIVERS)]
    if include_manager_visit:
        feature_names.append("manager_visit_worsening")
    feature_names.extend(("first_visit", "time_years"))
    feature_names.extend(
        f"restaurant:{level}" for level in restaurant_levels[1:]
    )
    feature_names.extend(
        f"questionnaire_version:{level}" for level in questionnaire_levels[1:]
    )
    return DesignSpec(
        restaurant_levels=restaurant_levels,
        questionnaire_levels=questionnaire_levels,
        first_week=min(item.week_start for item in responses),
        feature_names=tuple(feature_names),
    )


def design_matrix(
    responses: Sequence[Response],
    spec: DesignSpec,
    outcome: str = "low_overall",
    include_manager_visit: bool = False,
) -> tuple[np.ndarray, np.ndarray]:
    rows: list[list[float]] = []
    outcomes: list[int] = []
    restaurant_levels = spec.restaurant_levels[1:]
    questionnaire_levels = spec.questionnaire_levels[1:]
    for response in responses:
        if not response.model_usable_for(outcome, include_manager_visit):
            continue
        row = [1.0]
        row.extend(
            5.0 - float(getattr(response, driver))
            for driver in DRIVERS
        )
        if include_manager_visit:
            row.append(5.0 - float(response.manager_visit))
        row.append(1.0 if response.first_visit else 0.0)
        time_years = (response.week_start - spec.first_week).days / 364.0
        row.append(time_years)
        row.extend(
            1.0 if response.restaurant_id == level else 0.0
            for level in restaurant_levels
        )
        row.extend(
            1.0 if response.questionnaire_version == level else 0.0
            for level in questionnaire_levels
        )
        rows.append(row)
        outcomes.append(int(response.outcome_value(outcome)))
    return np.asarray(rows, dtype=np.float64), np.asarray(outcomes, dtype=np.float64)


def sigmoid(values: np.ndarray) -> np.ndarray:
    clipped = np.clip(values, -35.0, 35.0)
    return 1.0 / (1.0 + np.exp(-clipped))


def penalized_objective(
    x: np.ndarray,
    y: np.ndarray,
    beta: np.ndarray,
    c_value: float,
) -> float:
    logits = x @ beta
    likelihood = np.logaddexp(0.0, logits).sum() - y @ logits
    penalty = 0.5 * (1.0 / c_value) * float(beta[1:] @ beta[1:])
    return float(likelihood + penalty)


def fit_logistic(
    x: np.ndarray,
    y: np.ndarray,
    c_value: float,
    maximum_iterations: int = 100,
    tolerance: float = 1e-8,
) -> tuple[np.ndarray, bool, int]:
    if x.ndim != 2 or y.ndim != 1 or x.shape[0] != y.shape[0]:
        raise ValueError("Invalid model matrix dimensions.")
    if x.shape[0] == 0 or np.unique(y).size < 2:
        raise ValueError("Logistic regression requires both outcome classes.")

    beta = np.zeros(x.shape[1], dtype=np.float64)
    penalty_diagonal = np.ones(x.shape[1], dtype=np.float64) / c_value
    penalty_diagonal[0] = 0.0
    objective = penalized_objective(x, y, beta, c_value)

    for iteration in range(1, maximum_iterations + 1):
        probabilities = sigmoid(x @ beta)
        weights = np.clip(probabilities * (1.0 - probabilities), 1e-8, None)
        gradient = x.T @ (probabilities - y) + penalty_diagonal * beta
        hessian = x.T @ (x * weights[:, None])
        hessian.flat[:: hessian.shape[0] + 1] += penalty_diagonal
        hessian.flat[:: hessian.shape[0] + 1] += 1e-10
        try:
            step_direction = np.linalg.solve(hessian, gradient)
        except np.linalg.LinAlgError:
            step_direction = np.linalg.lstsq(hessian, gradient, rcond=None)[0]

        step_size = 1.0
        accepted = False
        candidate = beta
        candidate_objective = objective
        for _ in range(25):
            candidate = beta - step_size * step_direction
            candidate_objective = penalized_objective(
                x,
                y,
                candidate,
                c_value,
            )
            if candidate_objective <= objective + 1e-12:
                accepted = True
                break
            step_size *= 0.5
        if not accepted:
            return beta, False, iteration

        applied_step = candidate - beta
        beta = candidate
        objective = candidate_objective
        if np.max(np.abs(applied_step)) < tolerance:
            return beta, True, iteration
    return beta, False, maximum_iterations


def brier_score(y: np.ndarray, probabilities: np.ndarray) -> float:
    return float(np.mean((probabilities - y) ** 2))


def auc_score(y: np.ndarray, probabilities: np.ndarray) -> float | None:
    positives = int(y.sum())
    negatives = len(y) - positives
    if positives == 0 or negatives == 0:
        return None
    order = np.argsort(probabilities, kind="mergesort")
    ranks = np.empty(len(probabilities), dtype=np.float64)
    index = 0
    while index < len(probabilities):
        end = index + 1
        while (
            end < len(probabilities)
            and probabilities[order[end]] == probabilities[order[index]]
        ):
            end += 1
        average_rank = (index + 1 + end) / 2.0
        ranks[order[index:end]] = average_rank
        index = end
    positive_rank_sum = float(ranks[y == 1].sum())
    return (
        positive_rank_sum - positives * (positives + 1) / 2
    ) / (positives * negatives)


def fit_calibration(
    y: np.ndarray,
    probabilities: np.ndarray,
) -> tuple[float | None, float | None, bool]:
    if np.unique(y).size < 2:
        return None, None, False
    clipped = np.clip(probabilities, 1e-6, 1 - 1e-6)
    logits = np.log(clipped / (1.0 - clipped))
    x = np.column_stack((np.ones(len(logits)), logits))
    beta = np.array([0.0, 1.0], dtype=np.float64)
    for _ in range(100):
        fitted = sigmoid(x @ beta)
        weights = np.clip(fitted * (1.0 - fitted), 1e-8, None)
        gradient = x.T @ (fitted - y)
        hessian = x.T @ (x * weights[:, None])
        hessian.flat[::3] += 1e-10
        try:
            step = np.linalg.solve(hessian, gradient)
        except np.linalg.LinAlgError:
            return None, None, False
        beta -= step
        if np.max(np.abs(step)) < 1e-8:
            if np.all(np.isfinite(beta)):
                return float(beta[0]), float(beta[1]), True
            return None, None, False
    return None, None, False


def make_validation_folds(
    development: Sequence[Response],
    policy: dict[str, Any],
    outcome: str = "low_overall",
    include_manager_visit: bool = False,
) -> list[tuple[list[Response], list[Response], dict[str, Any]]]:
    weeks = sorted({item.week_start for item in development})
    minimum_training = policy["model"]["minimum_training_weeks_before_fold"]
    validation_weeks = policy["model"]["validation_block_weeks"]
    folds: list[tuple[list[Response], list[Response], dict[str, Any]]] = []
    start = minimum_training
    fold_number = 1
    while start + validation_weeks <= len(weeks):
        training_weeks = set(weeks[:start])
        validation_week_set = set(weeks[start : start + validation_weeks])
        training = [
            item for item in development if item.week_start in training_weeks
        ]
        validation = [
            item
            for item in development
            if item.week_start in validation_week_set
        ]
        usable_training = [
            item
            for item in training
            if item.model_usable_for(outcome, include_manager_visit)
        ]
        usable_validation = [
            item
            for item in validation
            if item.model_usable_for(outcome, include_manager_visit)
        ]
        if (
            usable_training
            and usable_validation
            and len(
                {item.outcome_value(outcome) for item in usable_training}
            )
            == 2
        ):
            folds.append(
                (
                    usable_training,
                    usable_validation,
                    {
                        "fold": fold_number,
                        "training_start_week": weeks[0].isoformat(),
                        "training_end_week": weeks[start - 1].isoformat(),
                        "validation_start_week": weeks[start].isoformat(),
                        "validation_end_week": weeks[
                            start + validation_weeks - 1
                        ].isoformat(),
                        "training_n": len(usable_training),
                        "validation_n": len(usable_validation),
                    },
                )
            )
            fold_number += 1
        start += validation_weeks
    return folds


def select_c_value(
    folds: Sequence[tuple[list[Response], list[Response], dict[str, Any]]],
    policy: dict[str, Any],
    outcome: str = "low_overall",
    include_manager_visit: bool = False,
) -> tuple[float | None, list[dict[str, Any]], list[dict[str, Any]]]:
    c_results: list[dict[str, Any]] = []
    fold_details = [metadata.copy() for _, _, metadata in folds]
    for c_value in policy["model"]["c_grid"]:
        squared_errors: list[np.ndarray] = []
        converged_folds = 0
        for training, validation, _ in folds:
            spec = make_design_spec(training, include_manager_visit)
            x_train, y_train = design_matrix(
                training,
                spec,
                outcome,
                include_manager_visit,
            )
            x_validation, y_validation = design_matrix(
                validation,
                spec,
                outcome,
                include_manager_visit,
            )
            beta, converged, _ = fit_logistic(
                x_train,
                y_train,
                float(c_value),
            )
            probabilities = sigmoid(x_validation @ beta)
            squared_errors.append((probabilities - y_validation) ** 2)
            converged_folds += int(converged)
        combined = np.concatenate(squared_errors)
        c_results.append(
            {
                "c": float(c_value),
                "brier_score": round_or_none(float(combined.mean()), 10),
                "folds": len(folds),
                "converged_folds": converged_folds,
                "all_folds_converged": converged_folds == len(folds),
            }
        )
    eligible_results = [
        result
        for result in c_results
        if (
            result["all_folds_converged"]
            or not policy["model"]["all_folds_must_converge"]
        )
    ]
    if not eligible_results:
        return None, c_results, fold_details
    best = min(
        eligible_results,
        key=lambda item: (item["brier_score"], item["c"]),
    )
    return float(best["c"]), c_results, fold_details


def bootstrap_feature_estimates(
    x: np.ndarray,
    y: np.ndarray,
    weeks: Sequence[date],
    c_value: float,
    policy: dict[str, Any],
    feature_indices: Sequence[int],
) -> tuple[np.ndarray, int]:
    requested = policy["model"]["bootstrap_replicates"]
    seed = policy["model"]["bootstrap_seed"]
    unique_weeks = sorted(set(weeks))
    indices_by_week = {
        week: np.asarray(
            [index for index, value in enumerate(weeks) if value == week],
            dtype=np.int64,
        )
        for week in unique_weeks
    }
    generator = np.random.default_rng(seed)
    estimates: list[np.ndarray] = []
    for _ in range(requested):
        sampled_weeks = generator.choice(
            np.asarray(unique_weeks, dtype=object),
            size=len(unique_weeks),
            replace=True,
        )
        sampled_indices = np.concatenate(
            [indices_by_week[value] for value in sampled_weeks]
        )
        bootstrap_y = y[sampled_indices]
        if np.unique(bootstrap_y).size < 2:
            continue
        bootstrap_x = x[sampled_indices]
        try:
            beta, converged, _ = fit_logistic(
                bootstrap_x,
                bootstrap_y,
                c_value,
            )
        except (ValueError, np.linalg.LinAlgError, FloatingPointError):
            continue
        if not converged:
            continue
        selected = beta[np.asarray(feature_indices, dtype=np.int64)]
        if np.all(np.isfinite(selected)):
            estimates.append(selected)
    if not estimates:
        return np.empty((0, len(feature_indices))), 0
    return np.vstack(estimates), len(estimates)


def make_store_gate(
    responses: Sequence[Response],
    policy: dict[str, Any],
) -> list[dict[str, Any]]:
    result = []
    gate = policy["shadow_gate"]
    for restaurant_id in sorted({item.restaurant_id for item in responses}):
        usable = [
            item
            for item in responses
            if item.restaurant_id == restaurant_id and item.model_usable
        ]
        events = sum(int(item.low_overall) for item in usable)
        eligible = (
            len(usable) >= gate["store_minimum_usable_responses"]
            and events >= gate["store_minimum_primary_events"]
        )
        result.append(
            {
                "restaurant_id": restaurant_id,
                "usable_responses": len(usable),
                "primary_events": events,
                "eligible": eligible,
                "suppressed": not eligible,
                "reason": (
                    None
                    if eligible
                    else "Store-specific estimates require at least "
                    f"{gate['store_minimum_usable_responses']} usable responses "
                    f"and {gate['store_minimum_primary_events']} primary events."
                ),
            }
        )
    return result


def make_input_manifest(
    payload: dict[str, Any],
    responses: Sequence[Response],
    totals: Sequence[PopulationTotal],
    policy: dict[str, Any],
    cycle_evidence: dict[str, Any] | None = None,
) -> dict[str, Any]:
    weeks = sorted({item.week_start for item in responses})
    export_id = payload.get("export_id")
    return {
        "schema_version": MANIFEST_SCHEMA,
        "input_schema_version": INPUT_SCHEMA,
        "policy_version": policy["version"],
        "policy_sha256": policy["source_sha256"],
        "policy_path": policy["source_path"],
        "source_sha256": str(payload["source_sha256"]).lower(),
        "export_id_sha256": (
            sha256_text(export_id)
            if isinstance(export_id, str) and export_id
            else None
        ),
        "record_set_sha256": response_record_set_hash(responses, totals),
        "record_set_hash_coverage": [
            "responses",
            "population_totals",
        ],
        "source_hash_definition": (
            "sha256(canonical JSON payload with source_sha256 omitted)"
        ),
        "source_hash_verified": True,
        "reporting_week_definition": "Monday-Sunday using local visit_date",
        "response_count": len(responses),
        "unique_response_count": len({item.response_id for item in responses}),
        "restaurant_count": len({item.restaurant_id for item in responses}),
        "questionnaire_version_count": len(
            {item.questionnaire_version for item in responses}
        ),
        "first_response_date": min(item.visit_date for item in responses).isoformat(),
        "last_response_date": max(item.visit_date for item in responses).isoformat(),
        "first_response_week": weeks[0].isoformat(),
        "last_response_week": weeks[-1].isoformat(),
        "distinct_response_weeks": len(weeks),
        "contiguous_response_weeks": longest_contiguous_week_run(weeks),
        "population_total_count": len(totals),
        "shadow_cycle_evidence": cycle_evidence
        or {
            "reported_completed_cycles": payload.get(
                "shadow_cycles_completed",
                0,
            ),
            "verified_completed_cycles": 0,
            "provenance": "input_reported_unverified",
            "durable_history_verified": False,
            "trusted_for_shadow_readiness": False,
            "trusted_for_promotion": False,
        },
        "row_level_data_persisted": False,
        "privacy_allowlist_enforced": True,
    }


def make_block_artifacts(
    payload: dict[str, Any],
    block: DataBlock,
    policy: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any], list[dict[str, Any]], str]:
    source_hash = payload.get("source_sha256") if isinstance(payload, dict) else None
    safe_source_hash = (
        source_hash.lower()
        if isinstance(source_hash, str)
        and len(source_hash) == 64
        and all(character in "0123456789abcdefABCDEF" for character in source_hash)
        else None
    )
    manifest = {
        "schema_version": MANIFEST_SCHEMA,
        "input_schema_version": (
            payload.get("schema_version") if isinstance(payload, dict) else None
        ),
        "policy_version": policy["version"],
        "policy_sha256": policy["source_sha256"],
        "policy_path": policy["source_path"],
        "source_sha256": safe_source_hash,
        "row_level_data_persisted": False,
        "privacy_allowlist_enforced": True,
        "manifest_status": "DataBlocked",
    }
    summary = {
        "schema_version": SUMMARY_SCHEMA,
        "policy_version": policy["version"],
        "policy_sha256": policy["source_sha256"],
        "status": "DataBlocked",
        "mode": "Shadow",
        "primary_outcome": policy["primary_outcome"],
        "secondary_outcome": policy["secondary_outcome"],
        "data_blockers": [{"code": block.code, "message": block.message}],
        "reconciliation": {"status": "DataBlocked"},
        "alerts": [],
        "driver_model": {"status": "NotRun"},
        "promotion": {
            "criteria_met": False,
            "eligible": False,
            "requires_explicit_approval": True,
        },
    }
    diagnostics = {
        "schema_version": DIAGNOSTICS_SCHEMA,
        "policy_version": policy["version"],
        "policy_sha256": policy["source_sha256"],
        "status": "DataBlocked",
        "data_blockers": [{"code": block.code, "message": block.message}],
        "privacy": {
            "stdin_only": True,
            "row_level_data_persisted": False,
            "individual_predictions_produced": False,
            "comments_or_contact_fields_accepted": False,
        },
    }
    card = render_model_card(summary, manifest, diagnostics)
    return summary, diagnostics, manifest, [], card


def evaluate_promotion(
    manifest: dict[str, Any],
    usable_count: int,
    event_count: int,
    fold_count: int,
    performance: dict[str, Any],
    top_factor_sign_stability: float | None,
    shadow_cycles: int,
    policy: dict[str, Any],
) -> dict[str, Any]:
    promotion = policy["promotion"]
    checks = [
        {
            "criterion": "distinct_response_weeks",
            "actual": manifest["distinct_response_weeks"],
            "required": promotion["minimum_distinct_weeks"],
            "met": manifest["distinct_response_weeks"]
            >= promotion["minimum_distinct_weeks"],
        },
        {
            "criterion": "usable_responses",
            "actual": usable_count,
            "required": promotion["minimum_usable_responses"],
            "met": usable_count >= promotion["minimum_usable_responses"],
        },
        {
            "criterion": "primary_events",
            "actual": event_count,
            "required": promotion["minimum_primary_events"],
            "met": event_count >= promotion["minimum_primary_events"],
        },
        {
            "criterion": "validation_folds",
            "actual": fold_count,
            "required": promotion["minimum_validation_folds"],
            "met": fold_count >= promotion["minimum_validation_folds"],
        },
        {
            "criterion": "holdout_brier_improvement_fraction",
            "actual": performance.get("brier_improvement_fraction"),
            "required": promotion["minimum_brier_improvement_fraction"],
            "met": (
                performance.get("brier_improvement_fraction") is not None
                and performance["brier_improvement_fraction"]
                >= promotion["minimum_brier_improvement_fraction"]
            ),
        },
        {
            "criterion": "holdout_auc",
            "actual": performance.get("auc"),
            "required": promotion["minimum_auc"],
            "met": (
                performance.get("auc") is not None
                and performance["auc"] >= promotion["minimum_auc"]
            ),
        },
        {
            "criterion": "calibration_model_converged",
            "actual": performance.get("calibration_converged", False),
            "required": True,
            "met": performance.get("calibration_converged") is True,
        },
        {
            "criterion": "calibration_slope",
            "actual": performance.get("calibration_slope"),
            "required": [
                promotion["calibration_slope_minimum"],
                promotion["calibration_slope_maximum"],
            ],
            "met": (
                performance.get("calibration_slope") is not None
                and promotion["calibration_slope_minimum"]
                <= performance["calibration_slope"]
                <= promotion["calibration_slope_maximum"]
            ),
        },
        {
            "criterion": "absolute_calibration_intercept",
            "actual": (
                None
                if performance.get("calibration_intercept") is None
                else abs(performance["calibration_intercept"])
            ),
            "required_maximum": promotion[
                "maximum_absolute_calibration_intercept"
            ],
            "met": (
                performance.get("calibration_intercept") is not None
                and abs(performance["calibration_intercept"])
                <= promotion["maximum_absolute_calibration_intercept"]
            ),
        },
        {
            "criterion": "top_factor_sign_stability",
            "actual": top_factor_sign_stability,
            "required": promotion["minimum_top_factor_sign_stability"],
            "met": (
                top_factor_sign_stability is not None
                and top_factor_sign_stability
                >= promotion["minimum_top_factor_sign_stability"]
            ),
        },
        {
            "criterion": "distinct_shadow_cycles",
            "actual": shadow_cycles,
            "required": promotion["minimum_shadow_cycles"],
            "met": shadow_cycles >= promotion["minimum_shadow_cycles"],
        },
    ]
    criteria_met = all(check["met"] for check in checks)
    return {
        "criteria": checks,
        "criteria_met": criteria_met,
        "eligible": False,
        "requires_explicit_approval": True,
        "reason": (
            "All numeric criteria are met, but promotion remains approval-gated."
            if criteria_met
            else "One or more promotion criteria are not yet met."
        ),
    }


def fit_outcome_model(
    responses: Sequence[Response],
    outcome: str,
    cycle_evidence: dict[str, Any],
    policy: dict[str, Any],
    include_manager_visit: bool = False,
) -> dict[str, Any]:
    if outcome not in OUTCOMES:
        raise ValueError(f"Unsupported outcome: {outcome}")
    analysis = (
        "manager_visit_sensitivity"
        if include_manager_visit
        else "driver_model"
    )
    usable = [
        item
        for item in responses
        if item.model_usable_for(outcome, include_manager_visit)
    ]
    events = sum(int(item.outcome_value(outcome)) for item in usable)
    weeks = sorted({item.week_start for item in usable})
    verified_cycles = int(cycle_evidence["verified_completed_cycles"])
    durable_history_verified = bool(
        cycle_evidence["durable_history_verified"]
    )
    gate_policy = policy["shadow_gate"]
    gate_checks = [
        {
            "criterion": "usable_responses",
            "actual": len(usable),
            "required": gate_policy["minimum_usable_responses"],
            "met": len(usable) >= gate_policy["minimum_usable_responses"],
        },
        {
            "criterion": "outcome_events",
            "actual": events,
            "required": gate_policy["minimum_primary_events"],
            "met": events >= gate_policy["minimum_primary_events"],
        },
        {
            "criterion": "distinct_response_weeks",
            "actual": len(weeks),
            "required": gate_policy["minimum_distinct_weeks"],
            "met": len(weeks) >= gate_policy["minimum_distinct_weeks"],
        },
        {
            "criterion": "distinct_shadow_cycles",
            "actual": verified_cycles,
            "required": gate_policy["minimum_cycles"],
            "met": verified_cycles >= gate_policy["minimum_cycles"],
        },
        {
            "criterion": "durable_shadow_cycle_history",
            "actual": cycle_evidence["provenance"],
            "required": "verified_append_only_aggregate_ledger",
            "met": durable_history_verified,
        },
    ]
    gate = {
        "passed": all(check["met"] for check in gate_checks),
        "criteria": gate_checks,
    }
    base_model = {
        "outcome": OUTCOMES[outcome]["csv_name"],
        "analysis": analysis,
        "gate": gate,
        "usable_responses": len(usable),
        "outcome_events": events,
        "distinct_response_weeks": len(weeks),
        "individual_predictions_produced": False,
    }
    if not gate["passed"]:
        sample_gate_passed = all(
            check["met"] for check in gate_checks[:3]
        )
        unverified_only = (
            sample_gate_passed and not durable_history_verified
        )
        gate_status = (
            "ShadowUnverified"
            if unverified_only
            else "ShadowSuppressed"
        )
        return {
            "model": {
                **base_model,
                "status": gate_status,
                "reason": (
                    "Durable aggregate shadow-cycle history is not verified."
                    if unverified_only
                    else "The outcome-specific shadow gate is not met."
                ),
            },
            "diagnostics": {
                "outcome": OUTCOMES[outcome]["csv_name"],
                "analysis": analysis,
                "status": gate_status,
                "gate": gate,
            },
            "estimates": [],
            "promotion_inputs": None,
        }

    holdout_week_count = policy["model"]["holdout_weeks"]
    holdout_start_week = weeks[-1] - timedelta(weeks=holdout_week_count - 1)
    development = [
        item for item in usable if item.week_start < holdout_start_week
    ]
    holdout = [
        item for item in usable if item.week_start >= holdout_start_week
    ]
    folds = make_validation_folds(
        development,
        policy,
        outcome,
        include_manager_visit,
    )
    if not folds:
        return {
            "model": {
                **base_model,
                "status": "ShadowSuppressed",
                "reason": "No valid expanding 4-week validation fold was available.",
            },
            "diagnostics": {
                "outcome": OUTCOMES[outcome]["csv_name"],
                "analysis": analysis,
                "status": "ShadowSuppressed",
                "gate": gate,
                "validation_fold_count": 0,
            },
            "estimates": [],
            "promotion_inputs": None,
        }

    selected_c, c_results, fold_details = select_c_value(
        folds,
        policy,
        outcome,
        include_manager_visit,
    )
    if selected_c is None:
        return {
            "model": {
                **base_model,
                "status": "ShadowSuppressed",
                "reason": "No regularization candidate converged in every validation fold.",
                "validation_fold_count": len(folds),
            },
            "diagnostics": {
                "outcome": OUTCOMES[outcome]["csv_name"],
                "analysis": analysis,
                "status": "ShadowSuppressed",
                "gate": gate,
                "validation": {
                    "method": "expanding origin with non-overlapping 4-week validation blocks",
                    "folds": fold_details,
                    "c_grid_results": c_results,
                    "selected_c": None,
                    "all_required_folds_converged": False,
                },
            },
            "estimates": [],
            "promotion_inputs": None,
        }
    spec = make_design_spec(development, include_manager_visit)
    x_development, y_development = design_matrix(
        development,
        spec,
        outcome,
        include_manager_visit,
    )
    x_holdout, y_holdout = design_matrix(
        holdout,
        spec,
        outcome,
        include_manager_visit,
    )
    beta, converged, iterations = fit_logistic(
        x_development,
        y_development,
        selected_c,
    )
    holdout_probabilities = sigmoid(x_holdout @ beta)
    model_brier = brier_score(y_holdout, holdout_probabilities)
    baseline_probability = float(y_development.mean())
    baseline_brier = brier_score(
        y_holdout,
        np.full(len(y_holdout), baseline_probability),
    )
    brier_improvement = (
        None
        if baseline_brier <= 0
        else (baseline_brier - model_brier) / baseline_brier
    )
    auc = auc_score(y_holdout, holdout_probabilities)
    (
        calibration_intercept,
        calibration_slope,
        calibration_converged,
    ) = fit_calibration(
        y_holdout,
        holdout_probabilities,
    )
    performance = {
        "holdout_start_week": min(item.week_start for item in holdout).isoformat(),
        "holdout_end_week": max(item.week_start for item in holdout).isoformat(),
        "holdout_n": len(holdout),
        "holdout_events": int(y_holdout.sum()),
        "brier_score": round_or_none(model_brier, 10),
        "baseline_brier_score": round_or_none(baseline_brier, 10),
        "brier_improvement_fraction": round_or_none(brier_improvement, 10),
        "auc": round_or_none(auc, 10),
        "calibration_intercept": round_or_none(calibration_intercept, 10),
        "calibration_slope": round_or_none(calibration_slope, 10),
        "calibration_converged": calibration_converged,
    }

    if include_manager_visit:
        reported_features = (("manager_visit", "Manager Visit", 1 + len(DRIVERS)),)
    else:
        reported_features = tuple(
            (driver, DRIVER_LABELS[driver], 1 + index)
            for index, driver in enumerate(DRIVERS)
        )
    feature_indices = [item[2] for item in reported_features]
    development_weeks = [item.week_start for item in development]
    bootstrap_values, completed_bootstraps = bootstrap_feature_estimates(
        x_development,
        y_development,
        development_weeks,
        selected_c,
        policy,
        feature_indices,
    )
    selected_c_result = next(
        item for item in c_results if item["c"] == selected_c
    )
    convergence_gate = {
        "selected_c_all_folds_converged": bool(
            selected_c_result["all_folds_converged"]
        ),
        "final_model_converged": bool(converged),
        "bootstrap_replicates_requested": policy["model"][
            "bootstrap_replicates"
        ],
        "bootstrap_replicates_completed": completed_bootstraps,
        "all_bootstrap_replicates_completed": (
            completed_bootstraps == policy["model"]["bootstrap_replicates"]
        ),
    }
    convergence_gate["passed"] = (
        (
            convergence_gate["selected_c_all_folds_converged"]
            or not policy["model"]["all_folds_must_converge"]
        )
        and (
            convergence_gate["final_model_converged"]
            or not policy["model"]["final_model_must_converge"]
        )
        and (
            convergence_gate["all_bootstrap_replicates_completed"]
            or not policy["model"]["all_bootstrap_repetitions_required"]
        )
    )
    if not convergence_gate["passed"]:
        return {
            "model": {
                **base_model,
                "status": "ShadowSuppressed",
                "reason": "Model convergence or bootstrap completeness requirements were not met.",
                "selected_c": selected_c,
                "validation_fold_count": len(folds),
                "convergence": convergence_gate,
            },
            "diagnostics": {
                "outcome": OUTCOMES[outcome]["csv_name"],
                "analysis": analysis,
                "status": "ShadowSuppressed",
                "gate": gate,
                "validation": {
                    "method": "expanding origin with non-overlapping 4-week validation blocks",
                    "folds": fold_details,
                    "c_grid_results": c_results,
                    "selected_c": selected_c,
                },
                "final_model": {
                    "converged": converged,
                    "iterations": iterations,
                },
                "bootstrap": {
                    "method": "response-week block bootstrap",
                    "seed": policy["model"]["bootstrap_seed"],
                    "requested_replicates": policy["model"][
                        "bootstrap_replicates"
                    ],
                    "completed_replicates": completed_bootstraps,
                },
                "convergence": convergence_gate,
            },
            "estimates": [],
            "promotion_inputs": None,
        }
    estimates: list[dict[str, Any]] = []
    for bootstrap_index, (_, label, feature_index) in enumerate(reported_features):
        coefficient = float(beta[feature_index])
        if completed_bootstraps:
            lower, upper = np.percentile(
                bootstrap_values[:, bootstrap_index],
                [2.5, 97.5],
            )
            if coefficient > 0:
                sign_stability = float(
                    np.mean(bootstrap_values[:, bootstrap_index] > 0)
                )
            elif coefficient < 0:
                sign_stability = float(
                    np.mean(bootstrap_values[:, bootstrap_index] < 0)
                )
            else:
                sign_stability = 0.0
        else:
            lower = math.nan
            upper = math.nan
            sign_stability = math.nan
        estimates.append(
            {
                "outcome": OUTCOMES[outcome]["csv_name"],
                "analysis": analysis,
                "factor": label,
                "coefficient_per_one_point_worsening": round_or_none(
                    coefficient,
                    10,
                ),
                "odds_ratio_per_one_point_worsening": round_or_none(
                    math.exp(coefficient),
                    10,
                ),
                "bootstrap_95_ci_coefficient_low": round_or_none(
                    float(lower),
                    10,
                ),
                "bootstrap_95_ci_coefficient_high": round_or_none(
                    float(upper),
                    10,
                ),
                "bootstrap_95_ci_odds_ratio_low": round_or_none(
                    math.exp(float(lower)) if math.isfinite(lower) else None,
                    10,
                ),
                "bootstrap_95_ci_odds_ratio_high": round_or_none(
                    math.exp(float(upper)) if math.isfinite(upper) else None,
                    10,
                ),
                "bootstrap_sign_stability": round_or_none(sign_stability, 10),
                "scope": "all_restaurants",
            }
        )

    top_factor = None
    if not include_manager_visit:
        top_estimate = max(
            estimates,
            key=lambda item: abs(item["coefficient_per_one_point_worsening"]),
        )
        top_factor = {
            "factor": top_estimate["factor"],
            "bootstrap_sign_stability": top_estimate[
                "bootstrap_sign_stability"
            ],
        }
    model = {
        **base_model,
        "status": "ShadowReady",
        "selected_c": selected_c,
        "selection_metric": "brier_score",
        "validation_fold_count": len(folds),
        "development_n": len(development),
        "development_events": int(y_development.sum()),
        "model_converged": converged,
        "model_iterations": iterations,
        "convergence": convergence_gate,
        "holdout": performance,
        "top_factor": top_factor,
        "reported_estimate_count": len(estimates),
    }
    diagnostics = {
        "outcome": OUTCOMES[outcome]["csv_name"],
        "analysis": analysis,
        "status": "ShadowReady",
        "gate": gate,
        "preprocessing": {
            "total_responses": len(responses),
            "model_usable_responses": len(usable),
            "excluded_incomplete_responses": len(responses) - len(usable),
            "feature_count_including_intercept": len(spec.feature_names),
            "restaurant_control_levels": len(spec.restaurant_levels),
            "questionnaire_version_control_levels": len(
                spec.questionnaire_levels
            ),
            "driver_coding": "5 minus rating; coefficient is per one-point worsening",
            "manager_visit_included": include_manager_visit,
            "time_control": "years since first development response week",
        },
        "validation": {
            "method": "expanding origin with non-overlapping 4-week validation blocks",
            "holdout_weeks": holdout_week_count,
            "folds": fold_details,
            "c_grid_results": c_results,
            "selected_c": selected_c,
        },
        "holdout": performance,
        "bootstrap": {
            "method": "response-week block bootstrap",
            "seed": policy["model"]["bootstrap_seed"],
            "requested_replicates": policy["model"]["bootstrap_replicates"],
            "completed_replicates": completed_bootstraps,
        },
        "convergence": convergence_gate,
    }
    return {
        "model": model,
        "diagnostics": diagnostics,
        "estimates": estimates,
        "promotion_inputs": {
            "usable_count": len(usable),
            "event_count": events,
            "fold_count": len(folds),
            "performance": performance,
            "top_factor_sign_stability": (
                None
                if top_factor is None
                else top_factor["bootstrap_sign_stability"]
            ),
        },
    }


def run_inference(
    payload: dict[str, Any],
    responses: Sequence[Response],
    totals: Sequence[PopulationTotal],
    reconciliation: dict[str, Any],
    policy: dict[str, Any],
    cycle_evidence: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any], list[dict[str, Any]], str]:
    manifest = make_input_manifest(
        payload,
        responses,
        totals,
        policy,
        cycle_evidence,
    )
    shadow_cycles = payload.get("shadow_cycles_completed", 0)
    if (
        isinstance(shadow_cycles, bool)
        or not isinstance(shadow_cycles, int)
        or shadow_cycles < 0
    ):
        raise DataBlock(
            "invalid_shadow_cycles",
            "shadow_cycles_completed must be a non-negative integer.",
        )

    outcome_results = {
        outcome: fit_outcome_model(
            responses,
            outcome,
            cycle_evidence,
            policy,
        )
        for outcome in OUTCOMES
    }
    manager_visit_provided = any(
        item.manager_visit is not None for item in responses
    )
    if manager_visit_provided:
        manager_sensitivity_result = fit_outcome_model(
            responses,
            "low_overall",
            cycle_evidence,
            policy,
            include_manager_visit=True,
        )
    else:
        manager_sensitivity_result = {
            "model": {
                "outcome": OUTCOMES["low_overall"]["csv_name"],
                "analysis": "manager_visit_sensitivity",
                "status": "NotProvided",
                "reason": "No conditionally eligible Manager Visit ratings were provided.",
                "individual_predictions_produced": False,
            },
            "diagnostics": {
                "outcome": OUTCOMES["low_overall"]["csv_name"],
                "analysis": "manager_visit_sensitivity",
                "status": "NotProvided",
            },
            "estimates": [],
            "promotion_inputs": None,
        }

    primary_result = outcome_results["low_overall"]
    secondary_result = outcome_results["recommend_detractor"]
    core_models_ready = all(
        result["model"]["status"] == "ShadowReady"
        for result in outcome_results.values()
    )
    if core_models_ready:
        overall_status = "ShadowReady"
    elif all(
        result["model"]["status"] == "ShadowUnverified"
        for result in outcome_results.values()
    ):
        overall_status = "ShadowUnverified"
    else:
        overall_status = "ShadowSuppressed"
    alerts = build_alerts(responses, reconciliation, policy)
    store_gate = make_store_gate(responses, policy)
    primary_usable = [
        item for item in responses if item.model_usable_for("low_overall")
    ]
    secondary_usable = [
        item
        for item in responses
        if item.model_usable_for("recommend_detractor")
    ]

    forecasting = {
        "implemented": False,
        "contiguous_response_weeks": manifest["contiguous_response_weeks"],
        "trend_data_threshold_met": (
            manifest["contiguous_response_weeks"]
            >= policy["forecasting"]["minimum_contiguous_weeks"]
        ),
        "seasonality_data_threshold_met": (
            manifest["contiguous_response_weeks"]
            >= policy["forecasting"]["minimum_seasonality_weeks"]
        ),
        "reason": "Forecasting is deferred and is not produced by the shadow model.",
    }
    combined_gate_criteria = []
    for outcome, result in outcome_results.items():
        for check in result["model"]["gate"]["criteria"]:
            combined_gate_criteria.append(
                {
                    **check,
                    "outcome": OUTCOMES[outcome]["csv_name"],
                }
            )
    shadow_gate = {
        "passed": core_models_ready,
        "criteria": combined_gate_criteria,
    }

    promotion_inputs = primary_result["promotion_inputs"]
    if promotion_inputs is None:
        promotion = {
            "criteria_met": False,
            "eligible": False,
            "requires_explicit_approval": True,
            "reason": "The primary shadow model is not ready.",
        }
    else:
        promotion = evaluate_promotion(
            manifest,
            promotion_inputs["usable_count"],
            promotion_inputs["event_count"],
            promotion_inputs["fold_count"],
            promotion_inputs["performance"],
            promotion_inputs["top_factor_sign_stability"],
            int(cycle_evidence["verified_completed_cycles"]),
            policy,
        )
        secondary_check = {
            "criterion": "secondary_outcome_model_ready",
            "actual": secondary_result["model"]["status"],
            "required": "ShadowReady",
            "met": secondary_result["model"]["status"] == "ShadowReady",
        }
        promotion["criteria"].append(secondary_check)
        promotion["criteria_met"] = (
            promotion["criteria_met"] and secondary_check["met"]
        )
        cycle_evidence_check = {
            "criterion": "verified_durable_shadow_cycle_evidence",
            "actual": cycle_evidence["provenance"],
            "required": "verified_durable_artifact_history",
            "met": bool(cycle_evidence["trusted_for_promotion"]),
        }
        promotion["criteria"].append(cycle_evidence_check)
        promotion["criteria_met"] = (
            promotion["criteria_met"] and cycle_evidence_check["met"]
        )
        if not promotion["criteria_met"]:
            promotion["reason"] = "One or more promotion criteria are not yet met."

    outcome_models = {
        outcome: result["model"] for outcome, result in outcome_results.items()
    }
    estimates = [
        estimate
        for result in outcome_results.values()
        for estimate in result["estimates"]
    ]
    estimates.extend(manager_sensitivity_result["estimates"])
    summary = {
        "schema_version": SUMMARY_SCHEMA,
        "policy_version": policy["version"],
        "policy_sha256": policy["source_sha256"],
        "status": overall_status,
        "mode": "Shadow",
        "primary_outcome": policy["primary_outcome"],
        "secondary_outcome": policy["secondary_outcome"],
        "reconciliation": reconciliation,
        "sample": {
            "responses": len(responses),
            "usable_responses": len(primary_usable),
            "primary_events": sum(
                int(item.low_overall) for item in primary_usable
            ),
            "secondary_usable_responses": len(secondary_usable),
            "secondary_events": sum(
                int(item.recommend_detractor) for item in secondary_usable
            ),
            "distinct_weeks": len(
                {item.week_start for item in primary_usable}
            ),
            "shadow_cycles_completed": int(
                cycle_evidence["verified_completed_cycles"]
            ),
        },
        "shadow_gate": shadow_gate,
        "store_specific": store_gate,
        "alerts": alerts,
        "forecasting": forecasting,
        "outcome_models": outcome_models,
        "driver_model": primary_result["model"],
        "manager_visit_sensitivity": manager_sensitivity_result["model"],
        "shadow_cycle_evidence": cycle_evidence,
        "promotion": promotion,
    }
    diagnostics = {
        "schema_version": DIAGNOSTICS_SCHEMA,
        "policy_version": policy["version"],
        "policy_sha256": policy["source_sha256"],
        "status": overall_status,
        "gate": shadow_gate,
        "models": {
            outcome: result["diagnostics"]
            for outcome, result in outcome_results.items()
        },
        "model": primary_result["diagnostics"],
        "manager_visit_sensitivity": manager_sensitivity_result[
            "diagnostics"
        ],
        "privacy": {
            "stdin_only": True,
            "row_level_data_persisted": False,
            "individual_predictions_produced": False,
            "comments_or_contact_fields_accepted": False,
            "reported_coefficients": list(DRIVER_LABELS.values()),
            "recommend_used_as_predictor": False,
            "manager_visit_used_in_primary_ranking": False,
        },
        "runtime": {
            "python": ".".join(map(str, sys.version_info[:3])),
            "packages": managed_runtime_versions(),
        },
    }
    card = render_model_card(summary, manifest, diagnostics)
    return summary, diagnostics, manifest, estimates, card


def render_model_card(
    summary: dict[str, Any],
    manifest: dict[str, Any],
    diagnostics: dict[str, Any],
) -> str:
    status = summary.get("status", "DataBlocked")
    response_count = manifest.get("response_count", "not available")
    distinct_weeks = manifest.get("distinct_response_weeks", "not available")
    blockers = summary.get("data_blockers", [])
    blocker_text = (
        "\n".join(
            f"- {item['code']}: {item['message']}" for item in blockers
        )
        if blockers
        else "- None."
    )
    outcome_models = summary.get("outcome_models", {})
    performance_parts = []
    for outcome in ("low_overall", "recommend_detractor"):
        model = outcome_models.get(outcome, {})
        holdout = model.get("holdout", {})
        label = OUTCOMES[outcome]["label"]
        if holdout:
            performance_parts.append(
                f"### {label}\n\n"
                f"- Status: {model.get('status')}\n"
                f"- Holdout Brier score: {holdout.get('brier_score')}\n"
                f"- Holdout AUC: {holdout.get('auc')}\n"
                f"- Calibration converged: {holdout.get('calibration_converged')}\n"
                f"- Calibration intercept: {holdout.get('calibration_intercept')}\n"
                f"- Calibration slope: {holdout.get('calibration_slope')}"
            )
        else:
            performance_parts.append(
                f"### {label}\n\n"
                f"- Status: {model.get('status', 'NotRun')}\n"
                "- Inferential performance was not calculated."
            )
    performance_lines = "\n\n".join(performance_parts)
    manager_status = summary.get("manager_visit_sensitivity", {}).get(
        "status",
        "NotProvided",
    )
    cycle_evidence = summary.get("shadow_cycle_evidence", {})
    return (
        "# GSS Shadow Driver Model Card\n\n"
        f"- Policy: `{summary.get('policy_version', EXPECTED_POLICY_VERSION)}`\n"
        f"- Status: **{status}**\n"
        f"- Mode: **Shadow**\n"
        f"- Primary model: **{outcome_models.get('low_overall', {}).get('status', 'NotRun')}**\n"
        f"- Secondary model: **{outcome_models.get('recommend_detractor', {}).get('status', 'NotRun')}**\n"
        f"- Manager Visit sensitivity: **{manager_status}**\n"
        f"- Responses in manifest: {response_count}\n"
        f"- Distinct response weeks: {distinct_weeks}\n\n"
        "## Intended use\n\n"
        "These models support aggregate operational learning about factors "
        "associated with low Overall (1-3) and Recommend detractor (0-6) "
        "outcomes. They do not score guests, make individual predictions, or "
        "authorize automated action. Recommend is an outcome in its own model "
        "and is never used as a predictor.\n\n"
        "## Data prerequisite\n\n"
        "The population export must reconcile by restaurant and non-overlapping "
        "13-week window before alerts or modeling run. Counts must be within "
        "max(1 response, 1%), score rates within 0.10 percentage points, stable "
        "response IDs must be unique, core score completeness must be at least "
        "99%, and visit dates and restaurant identifiers must be complete.\n\n"
        "### Data blockers\n\n"
        f"{blocker_text}\n\n"
        "## Method\n\n"
        "Separate L2-regularized logistic models for each approved outcome use "
        "Service, Culinary, Value, and Pace as one-point-worsening predictors. "
        "Controls are restaurant, First Visit, questionnaire version, and time. "
        "C is selected independently for each outcome from "
        "0.03, 0.1, 0.3, 1, and 3 using expanding 4-week validation and Brier "
        "score. The last 8 response weeks remain a temporal holdout. Driver "
        "uncertainty uses a 1,000-replicate response-week block bootstrap with "
        "seed 20260723 in production. Every selected validation fold, final "
        "fit, and requested bootstrap replicate must converge before an outcome "
        "is ShadowReady. Alerts use the Newcombe independent-proportions score "
        "interval and Benjamini-Hochberg false-discovery control.\n\n"
        "Manager Visit, when conditionally eligible, is fitted only in a "
        "separate primary-outcome sensitivity model and is excluded from the "
        "primary driver ranking.\n\n"
        "## Holdout performance\n\n"
        f"{performance_lines}\n\n"
        "## Shadow-cycle evidence\n\n"
        f"- Reported cycles: {cycle_evidence.get('reported_completed_cycles', 0)}\n"
        f"- Provenance: {cycle_evidence.get('provenance', 'not_available')}\n"
        "- Input-reported cycle counts can support shadow execution only. They "
        "are not trusted for promotion; verified durable artifact history is "
        "required.\n\n"
        "## Guardrails and limitations\n\n"
        "- Results remain shadow-only until every documented promotion criterion "
        "is met and a human explicitly approves promotion.\n"
        "- Store-specific results are suppressed below 500 usable responses and "
        "100 primary events per store.\n"
        "- Comments, names, contact fields, and free text are rejected by the "
        "input allowlist.\n"
        "- Raw response rows and individual predictions are not persisted.\n"
        "- Associations are not causal effects. Survey nonresponse and changes "
        "in questionnaire design may still bias estimates.\n"
        "- Forecasting is deferred; no forecast is produced by this model.\n"
    )


def write_atomic(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    file_descriptor, temporary_name = tempfile.mkstemp(
        prefix=".gss-model-",
        dir=str(path.parent),
    )
    try:
        with os.fdopen(file_descriptor, "w", encoding="utf-8", newline="") as handle:
            handle.write(content)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_name, path)
    finally:
        if os.path.exists(temporary_name):
            os.unlink(temporary_name)


@contextmanager
def cycle_ledger_lock(
    ledger_path: Path,
    timeout_seconds: float = 120.0,
) -> Iterator[None]:
    """Hold an OS-backed cross-process lock for one ledger transaction."""

    lock_path = ledger_path.with_name(ledger_path.name + ".lock")
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    with lock_path.open("a+b") as lock_handle:
        lock_handle.seek(0, os.SEEK_END)
        if lock_handle.tell() == 0:
            lock_handle.write(b"\0")
            lock_handle.flush()
            os.fsync(lock_handle.fileno())
        lock_handle.seek(0)
        if os.name == "nt":
            import msvcrt

            deadline = time.monotonic() + timeout_seconds
            while True:
                try:
                    msvcrt.locking(
                        lock_handle.fileno(),
                        msvcrt.LK_NBLCK,
                        1,
                    )
                    break
                except OSError as exc:
                    if time.monotonic() >= deadline:
                        raise TimeoutError(
                            "Timed out waiting for the aggregate shadow-cycle ledger lock."
                        ) from exc
                    time.sleep(0.05)
            try:
                yield
            finally:
                lock_handle.seek(0)
                msvcrt.locking(
                    lock_handle.fileno(),
                    msvcrt.LK_UNLCK,
                    1,
                )
        else:
            import fcntl

            fcntl.flock(lock_handle.fileno(), fcntl.LOCK_EX)
            try:
                yield
            finally:
                fcntl.flock(lock_handle.fileno(), fcntl.LOCK_UN)


def cycle_entry_sha256(entry: dict[str, Any]) -> str:
    hashable = {
        key: value for key, value in entry.items() if key != "entry_sha256"
    }
    return sha256_text(canonical_json(hashable))


def load_cycle_ledger(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {
            "schema_version": CYCLE_LEDGER_SCHEMA,
            "entries": [],
        }
    try:
        ledger = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise DataBlock(
            "invalid_cycle_ledger",
            "The aggregate shadow-cycle ledger could not be read.",
        ) from exc
    if (
        not isinstance(ledger, dict)
        or set(ledger) != {"schema_version", "entries"}
        or ledger.get("schema_version") != CYCLE_LEDGER_SCHEMA
        or not isinstance(ledger.get("entries"), list)
    ):
        raise DataBlock(
            "invalid_cycle_ledger",
            "The aggregate shadow-cycle ledger schema is invalid.",
        )
    previous_hash = ""
    allowed_statuses = {
        "ShadowReady",
        "ShadowSuppressed",
        "ShadowUnverified",
    }
    for entry in ledger["entries"]:
        if not isinstance(entry, dict):
            raise DataBlock(
                "invalid_cycle_ledger",
                "The aggregate shadow-cycle ledger contains an invalid entry.",
            )
        required = {
            "response_week",
            "source_sha256",
            "policy_sha256",
            "result_status",
            "previous_entry_sha256",
            "entry_sha256",
        }
        if set(entry) != required:
            raise DataBlock(
                "invalid_cycle_ledger",
                "The aggregate shadow-cycle ledger entry fields are invalid.",
            )
        try:
            response_week = date.fromisoformat(entry["response_week"])
        except (TypeError, ValueError) as exc:
            raise DataBlock(
                "invalid_cycle_ledger",
                "The aggregate shadow-cycle ledger week is invalid.",
            ) from exc
        hashes_valid = all(
            isinstance(entry[name], str)
            and len(entry[name]) == 64
            and all(
                character in "0123456789abcdef"
                for character in entry[name].lower()
            )
            for name in ("source_sha256", "policy_sha256", "entry_sha256")
        )
        if (
            response_week.weekday() != 0
            or not hashes_valid
            or entry["result_status"] not in allowed_statuses
            or entry["previous_entry_sha256"] != previous_hash
            or entry["entry_sha256"] != cycle_entry_sha256(entry)
        ):
            raise DataBlock(
                "invalid_cycle_ledger",
                "The aggregate shadow-cycle ledger integrity check failed.",
            )
        previous_hash = entry["entry_sha256"]
    return ledger


def get_cycle_evidence(
    ledger_path: Path | None,
    responses: Sequence[Response],
    source_sha256: str,
    policy: dict[str, Any],
    reported_cycles: int,
) -> dict[str, Any]:
    current_week = max(item.week_start for item in responses)
    if ledger_path is None:
        return {
            "reported_completed_cycles": reported_cycles,
            "verified_completed_cycles": 0,
            "current_cycle_week": current_week.isoformat(),
            "current_source_sha256": source_sha256,
            "provenance": "input_reported_unverified",
            "durable_history_verified": False,
            "trusted_for_shadow_readiness": False,
            "trusted_for_promotion": False,
            "ledger_sha256_before_run": None,
            "current_cycle_already_recorded": False,
        }

    ledger = load_cycle_ledger(ledger_path)
    matching_entries = [
        entry
        for entry in ledger["entries"]
        if entry["policy_sha256"] == policy["source_sha256"]
    ]
    completed_weeks = {
        entry["response_week"] for entry in matching_entries
    }
    current_already_recorded = any(
        entry["response_week"] == current_week.isoformat()
        and entry["source_sha256"] == source_sha256
        for entry in matching_entries
    )
    candidate_weeks = completed_weeks | {current_week.isoformat()}
    return {
        "reported_completed_cycles": reported_cycles,
        "verified_completed_cycles": len(candidate_weeks),
        "current_cycle_week": current_week.isoformat(),
        "current_source_sha256": source_sha256,
        "provenance": "verified_append_only_aggregate_ledger",
        "durable_history_verified": True,
        "trusted_for_shadow_readiness": True,
        "trusted_for_promotion": True,
        "ledger_sha256_before_run": sha256_text(canonical_json(ledger)),
        "current_cycle_already_recorded": current_already_recorded,
    }


def append_cycle_ledger(
    ledger_path: Path,
    evidence: dict[str, Any],
    result_status: str,
    policy: dict[str, Any],
) -> dict[str, Any]:
    ledger = load_cycle_ledger(ledger_path)
    already_recorded = any(
        entry["response_week"] == evidence["current_cycle_week"]
        and entry["source_sha256"] == evidence["current_source_sha256"]
        and entry["policy_sha256"] == policy["source_sha256"]
        for entry in ledger["entries"]
    )
    appended = False
    if not already_recorded:
        previous_hash = (
            ledger["entries"][-1]["entry_sha256"]
            if ledger["entries"]
            else ""
        )
        entry = {
            "response_week": evidence["current_cycle_week"],
            "source_sha256": evidence["current_source_sha256"],
            "policy_sha256": policy["source_sha256"],
            "result_status": result_status,
            "previous_entry_sha256": previous_hash,
        }
        entry["entry_sha256"] = cycle_entry_sha256(entry)
        ledger["entries"].append(entry)
        write_atomic(
            ledger_path,
            json.dumps(ledger, indent=2, sort_keys=True, ensure_ascii=True) + "\n",
        )
        appended = True
    return {
        **evidence,
        "current_cycle_appended": appended,
        "ledger_sha256_after_run": sha256_text(canonical_json(ledger)),
    }


def estimates_csv(estimates: Sequence[dict[str, Any]]) -> str:
    fields = [
        "outcome",
        "analysis",
        "factor",
        "coefficient_per_one_point_worsening",
        "odds_ratio_per_one_point_worsening",
        "bootstrap_95_ci_coefficient_low",
        "bootstrap_95_ci_coefficient_high",
        "bootstrap_95_ci_odds_ratio_low",
        "bootstrap_95_ci_odds_ratio_high",
        "bootstrap_sign_stability",
        "scope",
    ]
    stream = io.StringIO(newline="")
    writer = csv.DictWriter(stream, fieldnames=fields, lineterminator="\n")
    writer.writeheader()
    for estimate in estimates:
        writer.writerow(estimate)
    return stream.getvalue()


def write_artifacts(
    output_directory: Path,
    summary: dict[str, Any],
    diagnostics: dict[str, Any],
    manifest: dict[str, Any],
    estimates: Sequence[dict[str, Any]],
    model_card: str,
) -> None:
    output_directory.mkdir(parents=True, exist_ok=True)
    write_atomic(
        output_directory / "model_summary.json",
        json.dumps(summary, indent=2, sort_keys=True, ensure_ascii=True) + "\n",
    )
    write_atomic(
        output_directory / "model_diagnostics.json",
        json.dumps(diagnostics, indent=2, sort_keys=True, ensure_ascii=True) + "\n",
    )
    write_atomic(
        output_directory / "input_manifest.json",
        json.dumps(manifest, indent=2, sort_keys=True, ensure_ascii=True) + "\n",
    )
    write_atomic(
        output_directory / "model_estimates.csv",
        estimates_csv(estimates),
    )
    write_atomic(output_directory / "model_card.md", model_card)


def run_pipeline(
    payload: dict[str, Any],
    output_directory: Path,
    policy_override: dict[str, Any] | None = None,
    cycle_ledger_path: Path | None = None,
) -> dict[str, Any]:
    policy = policy_with_override(policy_override)
    try:
        responses, totals = parse_payload(payload)
        reconciliation = reconcile_population(responses, totals, policy)
        if reconciliation["status"] != "Passed":
            block = DataBlock(
                "population_reconciliation_failed",
                "The population export did not meet all reconciliation criteria.",
            )
            manifest = make_input_manifest(payload, responses, totals, policy)
            summary, diagnostics, _, estimates, card = make_block_artifacts(
                payload,
                block,
                policy,
            )
            summary["reconciliation"] = reconciliation
            diagnostics["reconciliation"] = reconciliation
            write_artifacts(
                output_directory,
                summary,
                diagnostics,
                manifest,
                estimates,
                card,
            )
            return summary
        ledger_transaction = (
            cycle_ledger_lock(cycle_ledger_path)
            if cycle_ledger_path is not None
            else nullcontext()
        )
        with ledger_transaction:
            cycle_evidence = get_cycle_evidence(
                cycle_ledger_path,
                responses,
                str(payload["source_sha256"]).lower(),
                policy,
                payload.get("shadow_cycles_completed", 0),
            )
            summary, diagnostics, manifest, estimates, card = run_inference(
                payload,
                responses,
                totals,
                reconciliation,
                policy,
                cycle_evidence,
            )
            if cycle_ledger_path is not None:
                updated_cycle_evidence = append_cycle_ledger(
                    cycle_ledger_path,
                    summary["shadow_cycle_evidence"],
                    summary["status"],
                    policy,
                )
                summary["shadow_cycle_evidence"] = updated_cycle_evidence
                manifest["shadow_cycle_evidence"] = updated_cycle_evidence
                diagnostics["shadow_cycle_evidence"] = updated_cycle_evidence
                card = render_model_card(summary, manifest, diagnostics)
    except DataBlock as block:
        summary, diagnostics, manifest, estimates, card = make_block_artifacts(
            payload,
            block,
            policy,
        )

    write_artifacts(
        output_directory,
        summary,
        diagnostics,
        manifest,
        estimates,
        card,
    )
    return summary


def parse_arguments(arguments: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run privacy-safe GSS statistical modeling from stdin.",
    )
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument(
        "--output-dir",
        help="Directory for aggregate model artifacts.",
    )
    mode.add_argument(
        "--compute-source-sha256",
        action="store_true",
        help=(
            "Read JSON from stdin and print the canonical SHA-256 with "
            "source_sha256 omitted; write no artifacts."
        ),
    )
    parser.add_argument(
        "--cycle-ledger",
        help="Append-only aggregate cycle ledger used for verified shadow readiness.",
    )
    return parser.parse_args(arguments)


def main(arguments: Sequence[str] | None = None) -> int:
    args = parse_arguments(arguments if arguments is not None else sys.argv[1:])
    if sys.version_info[:2] != (3, 12):
        print(
            "GSS shadow modeling requires managed Python 3.12.",
            file=sys.stderr,
        )
        return 2
    try:
        ensure_managed_runtime()
    except RuntimeError as error:
        print(str(error), file=sys.stderr)
        return 2
    try:
        raw_input = sys.stdin.read()
        if not raw_input.strip():
            raise DataBlock("empty_stdin", "A JSON population export is required on stdin.")
        try:
            payload = json.loads(raw_input)
        except json.JSONDecodeError as exc:
            raise DataBlock(
                "invalid_json",
                "The stdin payload is not valid JSON.",
            ) from exc
        if args.compute_source_sha256:
            print(canonical_source_sha256(payload))
            return 0
        summary = run_pipeline(
            payload,
            Path(args.output_dir),
            cycle_ledger_path=(
                Path(args.cycle_ledger) if args.cycle_ledger else None
            ),
        )
    except DataBlock as block:
        if args.compute_source_sha256:
            print(block.message, file=sys.stderr)
            return 2
        policy = policy_with_override(None)
        payload = {}
        summary, diagnostics, manifest, estimates, card = make_block_artifacts(
            payload,
            block,
            policy,
        )
        write_artifacts(
            Path(args.output_dir),
            summary,
            diagnostics,
            manifest,
            estimates,
            card,
        )
    except Exception:
        print(
            "GSS shadow modeling failed because of an internal runtime error; "
            "no row data was logged.",
            file=sys.stderr,
        )
        return 1

    print(
        canonical_json(
            {
                "schema_version": SUMMARY_SCHEMA,
                "status": summary["status"],
                "mode": "Shadow",
                "artifact_count": len(REQUIRED_ARTIFACTS),
            }
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
