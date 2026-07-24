"""Focused synthetic tests for the privacy-safe GSS shadow model."""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = REPO_ROOT / "scripts" / "gss_shadow_model.py"
SPEC = importlib.util.spec_from_file_location("gss_shadow_model", MODULE_PATH)
assert SPEC is not None and SPEC.loader is not None
MODEL = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODEL
SPEC.loader.exec_module(MODEL)


def make_payload(
    week_count: int = 40,
    responses_per_week: int = 20,
    manager_visit: bool = False,
    shadow_cycles: int = 8,
) -> dict:
    first_week = date(2025, 9, 8)
    responses = []
    for week_index in range(week_count):
        week_start = first_week + timedelta(weeks=week_index)
        for response_index in range(responses_per_week):
            ordinal = week_index * responses_per_week + response_index
            restaurant_id = "9354" if response_index % 2 == 0 else "9355"
            service = 1 + ((ordinal * 3 + week_index) % 5)
            culinary = 1 + ((ordinal * 2 + week_index) % 5)
            value = 1 + ((ordinal * 7 + week_index * 2) % 5)
            pace = 1 + ((ordinal * 11 + week_index) % 5)
            low = service <= 2 or (value == 1 and response_index % 3 == 0)
            overall = 2 if low else 5
            recommend = 4 if (low or ordinal % 11 == 0) else 9
            response = {
                    "response_id": f"RID-{ordinal:05d}",
                    "restaurant_id": restaurant_id,
                    "visit_date": (
                        week_start + timedelta(days=response_index % 7)
                    ).isoformat(),
                    "overall": overall,
                    "service": service,
                    "culinary": culinary,
                    "value": value,
                    "pace": pace,
                    "recommend": recommend,
                    "first_visit": ordinal % 5 == 0,
                    "questionnaire_version": (
                        "v1" if week_index < week_count // 2 else "v2"
                    ),
                    "conditional_eligibility": {
                        "manager_visit": manager_visit
                    },
                }
            if manager_visit:
                response["manager_visit"] = 1 + ((ordinal * 13) % 5)
            responses.append(response)

    max_week = first_week + timedelta(weeks=week_count - 1)
    bounds = MODEL.get_window_bounds(max_week)
    population_totals = []
    for restaurant_id in ("9354", "9355"):
        for window in MODEL.WINDOWS:
            start, end = bounds[window]
            rows = [
                row
                for row in responses
                if row["restaurant_id"] == restaurant_id
                and start
                <= MODEL.monday_week_start(
                    date.fromisoformat(row["visit_date"])
                )
                <= end
            ]
            low_rate = (
                sum(row["overall"] <= 3 for row in rows) / len(rows) * 100
                if rows
                else 0
            )
            detractor_rate = (
                sum(row["recommend"] <= 6 for row in rows) / len(rows) * 100
                if rows
                else 0
            )
            population_totals.append(
                {
                    "restaurant_id": restaurant_id,
                    "window": window,
                    "response_count": len(rows),
                    "low_overall_rate_pct": low_rate,
                    "recommend_detractor_rate_pct": detractor_rate,
                }
            )

    payload = {
        "schema_version": MODEL.INPUT_SCHEMA,
        "export_id": "synthetic-export",
        "shadow_cycles_completed": shadow_cycles,
        "responses": responses,
        "population_totals": population_totals,
    }
    payload["source_sha256"] = MODEL.canonical_source_sha256(payload)
    return payload


def refresh_source_hash(payload: dict) -> None:
    payload["source_sha256"] = MODEL.canonical_source_sha256(payload)


def latest_response_week(payload: dict) -> date:
    return max(
        MODEL.monday_week_start(date.fromisoformat(row["visit_date"]))
        for row in payload["responses"]
    )


def seed_cycle_ledger(
    ledger_path: Path,
    payload: dict,
    prior_cycle_count: int = 7,
) -> None:
    policy = MODEL.load_tracked_policy()
    current_week = latest_response_week(payload)
    entries = []
    previous_hash = ""
    for cycle_offset in range(prior_cycle_count, 0, -1):
        response_week = current_week - timedelta(weeks=cycle_offset)
        entry = {
            "response_week": response_week.isoformat(),
            "source_sha256": hashlib.sha256(
                f"synthetic-cycle-{response_week.isoformat()}".encode("utf-8")
            ).hexdigest(),
            "policy_sha256": policy["source_sha256"],
            "result_status": "ShadowSuppressed",
            "previous_entry_sha256": previous_hash,
        }
        entry["entry_sha256"] = MODEL.cycle_entry_sha256(entry)
        entries.append(entry)
        previous_hash = entry["entry_sha256"]
    ledger_path.parent.mkdir(parents=True, exist_ok=True)
    ledger_path.write_text(
        json.dumps(
            {
                "schema_version": MODEL.CYCLE_LEDGER_SCHEMA,
                "entries": entries,
            },
            indent=2,
            sort_keys=True,
        )
        + "\n",
        encoding="utf-8",
    )


def run_with_verified_history(
    payload: dict,
    root: Path,
    policy_override: dict | None = None,
) -> tuple[dict, Path, Path]:
    output = root / "artifacts"
    ledger = root / "state" / "shadow-model-cycle-ledger.json"
    seed_cycle_ledger(ledger, payload)
    summary = MODEL.run_pipeline(
        payload,
        output,
        policy_override,
        cycle_ledger_path=ledger,
    )
    return summary, output, ledger


def quote_powershell_literal(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


class ShadowModelTests(unittest.TestCase):
    def test_monday_sunday_reporting_week_boundaries(self) -> None:
        monday = date(2026, 7, 13)
        sunday = date(2026, 7, 19)
        next_monday = date(2026, 7, 20)
        self.assertEqual(MODEL.monday_week_start(monday), monday)
        self.assertEqual(MODEL.monday_week_start(sunday), monday)
        self.assertEqual(
            MODEL.monday_week_start(next_monday),
            next_monday,
        )

    def test_canonical_source_hash_covers_responses_and_population_totals(self) -> None:
        payload = make_payload()
        original_hash = payload["source_sha256"]
        without_hash = copy.deepcopy(payload)
        without_hash.pop("source_sha256")
        blank_hash = copy.deepcopy(payload)
        blank_hash["source_sha256"] = ""
        unrelated_hash = copy.deepcopy(payload)
        unrelated_hash["source_sha256"] = "f" * 64
        self.assertEqual(
            MODEL.canonical_source_sha256(without_hash),
            original_hash,
        )
        self.assertEqual(
            MODEL.canonical_source_sha256(blank_hash),
            original_hash,
        )
        self.assertEqual(
            MODEL.canonical_source_sha256(unrelated_hash),
            original_hash,
        )

        changed_response = copy.deepcopy(payload)
        changed_response["responses"][0]["service"] = 5
        self.assertNotEqual(
            MODEL.canonical_source_sha256(changed_response),
            original_hash,
        )
        changed_total = copy.deepcopy(payload)
        changed_total["population_totals"][0]["response_count"] += 1
        self.assertNotEqual(
            MODEL.canonical_source_sha256(changed_total),
            original_hash,
        )

    def test_mismatched_source_hash_is_data_blocked_without_cycle_append(self) -> None:
        payload = make_payload()
        payload["population_totals"][0]["response_count"] += 1
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            ledger = root / "state" / "shadow-model-cycle-ledger.json"
            summary = MODEL.run_pipeline(
                payload,
                root / "artifacts",
                {"bootstrap_replicates": 5},
                cycle_ledger_path=ledger,
            )
            self.assertEqual(summary["status"], "DataBlocked")
            self.assertEqual(
                summary["data_blockers"][0]["code"],
                "source_hash_mismatch",
            )
            self.assertFalse(ledger.exists())

    def test_reconciled_population_produces_only_aggregate_artifacts(self) -> None:
        payload = make_payload()
        with tempfile.TemporaryDirectory() as temporary_directory:
            summary, output, ledger = run_with_verified_history(
                payload,
                Path(temporary_directory),
                {"bootstrap_replicates": 20},
            )

            self.assertEqual(summary["status"], "ShadowReady")
            self.assertEqual(summary["reconciliation"]["status"], "Passed")
            self.assertEqual(
                sorted(path.name for path in output.iterdir()),
                sorted(MODEL.REQUIRED_ARTIFACTS),
            )
            self.assertEqual(
                summary["outcome_models"]["low_overall"][
                    "reported_estimate_count"
                ],
                4,
            )
            self.assertEqual(
                summary["outcome_models"]["recommend_detractor"][
                    "reported_estimate_count"
                ],
                4,
            )
            self.assertEqual(
                summary["outcome_models"]["recommend_detractor"]["status"],
                "ShadowReady",
            )
            self.assertTrue(
                all(item["suppressed"] for item in summary["store_specific"])
            )

            diagnostics = json.loads(
                (output / "model_diagnostics.json").read_text(encoding="utf-8")
            )
            manifest = json.loads(
                (output / "input_manifest.json").read_text(encoding="utf-8")
            )
            self.assertEqual(
                manifest["policy_version"],
                "gss-analysis-policy/v3",
            )
            self.assertEqual(
                manifest["policy_sha256"],
                hashlib.sha256(
                    (REPO_ROOT / "config" / "analysis-policy.json").read_bytes()
                ).hexdigest(),
            )
            self.assertTrue(manifest["source_hash_verified"])
            self.assertEqual(
                manifest["record_set_hash_coverage"],
                ["responses", "population_totals"],
            )
            self.assertEqual(
                manifest["shadow_cycle_evidence"]["verified_completed_cycles"],
                8,
            )
            self.assertTrue(
                manifest["shadow_cycle_evidence"][
                    "trusted_for_shadow_readiness"
                ]
            )
            self.assertEqual(
                diagnostics["models"]["low_overall"]["bootstrap"][
                    "requested_replicates"
                ],
                20,
            )
            self.assertEqual(
                diagnostics["models"]["low_overall"]["bootstrap"][
                    "completed_replicates"
                ],
                20,
            )
            self.assertEqual(
                diagnostics["models"]["recommend_detractor"]["bootstrap"][
                    "completed_replicates"
                ],
                20,
            )
            all_artifact_text = "\n".join(
                path.read_text(encoding="utf-8")
                for path in output.iterdir()
            )
            self.assertNotIn("RID-", all_artifact_text)
            self.assertNotIn("synthetic-export", all_artifact_text)
            self.assertNotIn("conditional_eligibility", all_artifact_text)
            estimates = (output / "model_estimates.csv").read_text(
                encoding="utf-8"
            )
            self.assertIn("low_overall_1_to_3,driver_model", estimates)
            self.assertIn(
                "recommend_detractor_0_to_6,driver_model",
                estimates,
            )
            self.assertEqual(len(estimates.strip().splitlines()), 9)
            ledger_text = ledger.read_text(encoding="utf-8")
            self.assertNotIn("RID-", ledger_text)
            self.assertNotIn("responses", ledger_text)

    def test_same_input_and_seed_are_byte_deterministic(self) -> None:
        payload = make_payload()
        with (
            tempfile.TemporaryDirectory() as first_directory,
            tempfile.TemporaryDirectory() as second_directory,
        ):
            _, first_output, _ = run_with_verified_history(
                payload,
                Path(first_directory),
                {"bootstrap_replicates": 12},
            )
            _, second_output, _ = run_with_verified_history(
                payload,
                Path(second_directory),
                {"bootstrap_replicates": 12},
            )
            for artifact in MODEL.REQUIRED_ARTIFACTS:
                self.assertEqual(
                    (first_output / artifact).read_bytes(),
                    (second_output / artifact).read_bytes(),
                    artifact,
                )

    def test_reconciliation_mismatch_blocks_inference(self) -> None:
        payload = make_payload()
        payload["population_totals"][0]["response_count"] += 20
        refresh_source_hash(payload)
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            output = root / "artifacts"
            ledger = root / "state" / "shadow-model-cycle-ledger.json"
            seed_cycle_ledger(ledger, payload)
            ledger_before = ledger.read_bytes()
            summary = MODEL.run_pipeline(
                payload,
                output,
                {"bootstrap_replicates": 5},
                cycle_ledger_path=ledger,
            )

            self.assertEqual(summary["status"], "DataBlocked")
            self.assertEqual(
                summary["reconciliation"]["status"],
                "DataBlocked",
            )
            self.assertEqual(summary["driver_model"]["status"], "NotRun")
            estimates = (output / "model_estimates.csv").read_text(
                encoding="utf-8"
            )
            self.assertEqual(len(estimates.strip().splitlines()), 1)
            self.assertEqual(ledger.read_bytes(), ledger_before)

    def test_unexpected_row_field_is_privacy_blocker_and_is_not_persisted(self) -> None:
        payload = make_payload()
        payload["responses"][0]["comment"] = "SECRET-GUEST-COMMENT"
        refresh_source_hash(payload)
        with tempfile.TemporaryDirectory() as temporary_directory:
            output = Path(temporary_directory)
            summary = MODEL.run_pipeline(
                payload,
                output,
                {"bootstrap_replicates": 5},
            )

            self.assertEqual(summary["status"], "DataBlocked")
            self.assertEqual(
                summary["data_blockers"][0]["code"],
                "privacy_unexpected_fields",
            )
            all_artifact_text = "\n".join(
                path.read_text(encoding="utf-8")
                for path in output.iterdir()
            )
            self.assertNotIn("SECRET-GUEST-COMMENT", all_artifact_text)
            self.assertNotIn("RID-", all_artifact_text)

    def test_below_shadow_minimum_is_suppressed_without_model(self) -> None:
        payload = make_payload(week_count=26, responses_per_week=4)
        with tempfile.TemporaryDirectory() as temporary_directory:
            summary = MODEL.run_pipeline(
                copy.deepcopy(payload),
                Path(temporary_directory),
                {"bootstrap_replicates": 5},
            )
            self.assertEqual(summary["status"], "ShadowSuppressed")
            self.assertFalse(summary["shadow_gate"]["passed"])
            self.assertEqual(
                summary["driver_model"]["status"],
                "ShadowSuppressed",
            )

    def test_newcombe_independent_score_interval_regression(self) -> None:
        lower, upper = MODEL.newcombe_independent_difference_interval(
            30,
            100,
            20,
            100,
        )
        self.assertAlmostEqual(lower, -0.0202493621530134, places=12)
        self.assertAlmostEqual(upper, 0.21673435401778587, places=12)
        endpoint_subtraction = (
            MODEL.wilson_interval(30, 100)[0]
            - MODEL.wilson_interval(20, 100)[1]
        )
        self.assertNotAlmostEqual(lower, endpoint_subtraction, places=6)

    def test_manager_visit_is_conditionally_eligible_sensitivity_only(self) -> None:
        payload = make_payload(manager_visit=True)
        with tempfile.TemporaryDirectory() as temporary_directory:
            summary, output, _ = run_with_verified_history(
                payload,
                Path(temporary_directory),
                {"bootstrap_replicates": 12},
            )
            self.assertEqual(summary["status"], "ShadowReady")
            self.assertEqual(
                summary["manager_visit_sensitivity"]["status"],
                "ShadowReady",
            )
            self.assertNotEqual(
                summary["driver_model"]["top_factor"]["factor"],
                "Manager Visit",
            )
            estimates = (output / "model_estimates.csv").read_text(
                encoding="utf-8"
            )
            self.assertIn(
                "low_overall_1_to_3,manager_visit_sensitivity,Manager Visit",
                estimates,
            )

    def test_manager_visit_without_eligibility_is_data_blocked(self) -> None:
        payload = make_payload()
        payload["responses"][0]["manager_visit"] = 4
        refresh_source_hash(payload)
        with tempfile.TemporaryDirectory() as temporary_directory:
            summary = MODEL.run_pipeline(
                payload,
                Path(temporary_directory),
                {"bootstrap_replicates": 5},
            )
            self.assertEqual(summary["status"], "DataBlocked")
            self.assertEqual(
                summary["data_blockers"][0]["code"],
                "manager_visit_not_eligible",
            )

    def test_self_reported_cycles_never_create_shadow_ready_status(self) -> None:
        payload = make_payload(shadow_cycles=999)
        with tempfile.TemporaryDirectory() as temporary_directory:
            output = Path(temporary_directory)
            summary = MODEL.run_pipeline(
                payload,
                output,
                {"bootstrap_replicates": 5},
            )
            self.assertEqual(summary["status"], "ShadowUnverified")
            self.assertFalse(summary["shadow_gate"]["passed"])
            self.assertFalse(summary["promotion"]["criteria_met"])
            self.assertFalse(summary["promotion"]["eligible"])
            self.assertEqual(
                summary["shadow_cycle_evidence"]["provenance"],
                "input_reported_unverified",
            )
            self.assertFalse(
                summary["shadow_cycle_evidence"][
                    "trusted_for_shadow_readiness"
                ]
            )
            self.assertFalse(
                summary["shadow_cycle_evidence"]["trusted_for_promotion"]
            )
            estimates = (output / "model_estimates.csv").read_text(
                encoding="utf-8"
            )
            self.assertEqual(len(estimates.strip().splitlines()), 1)

    def test_durable_cycle_ledger_is_idempotent_for_same_week_and_source(self) -> None:
        payload = make_payload()
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            ledger = root / "state" / "shadow-model-cycle-ledger.json"
            seed_cycle_ledger(ledger, payload)
            first = MODEL.run_pipeline(
                payload,
                root / "first",
                {"bootstrap_replicates": 5},
                cycle_ledger_path=ledger,
            )
            first_ledger = MODEL.load_cycle_ledger(ledger)
            second = MODEL.run_pipeline(
                payload,
                root / "second",
                {"bootstrap_replicates": 5},
                cycle_ledger_path=ledger,
            )
            second_ledger = MODEL.load_cycle_ledger(ledger)

            self.assertEqual(first["status"], "ShadowReady")
            self.assertTrue(
                first["shadow_cycle_evidence"]["current_cycle_appended"]
            )
            self.assertEqual(
                first["shadow_cycle_evidence"]["verified_completed_cycles"],
                8,
            )
            self.assertEqual(second["status"], "ShadowReady")
            self.assertFalse(
                second["shadow_cycle_evidence"]["current_cycle_appended"]
            )
            self.assertTrue(
                second["shadow_cycle_evidence"][
                    "current_cycle_already_recorded"
                ]
            )
            self.assertEqual(len(first_ledger["entries"]), 8)
            self.assertEqual(second_ledger, first_ledger)

    def test_concurrent_processes_preserve_both_cycle_entries(self) -> None:
        first_payload = make_payload(
            week_count=26,
            responses_per_week=4,
        )
        second_payload = make_payload(
            week_count=27,
            responses_per_week=4,
        )
        runner = r"""
import importlib.util
import json
import sys
import time
from pathlib import Path

module_path = Path(sys.argv[1])
ledger_path = Path(sys.argv[2])
output_path = Path(sys.argv[3])
spec = importlib.util.spec_from_file_location("gss_shadow_model_child", module_path)
module = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = module
spec.loader.exec_module(module)
original_write_atomic = module.write_atomic

def delayed_ledger_write(path, content):
    if Path(path) == ledger_path:
        time.sleep(0.75)
    original_write_atomic(path, content)

module.write_atomic = delayed_ledger_write
payload = json.loads(sys.stdin.read())
summary = module.run_pipeline(
    payload,
    output_path,
    {"bootstrap_replicates": 1},
    cycle_ledger_path=ledger_path,
)
print(summary["status"])
"""
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            ledger = root / "state" / "shadow-model-cycle-ledger.json"
            command_prefix = [
                sys.executable,
                "-c",
                runner,
                str(MODULE_PATH),
                str(ledger),
            ]
            first_process = subprocess.Popen(
                [*command_prefix, str(root / "first")],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            second_process = subprocess.Popen(
                [*command_prefix, str(root / "second")],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            assert first_process.stdin is not None
            assert second_process.stdin is not None
            first_process.stdin.write(json.dumps(first_payload))
            first_process.stdin.close()
            first_process.stdin = None
            second_process.stdin.write(json.dumps(second_payload))
            second_process.stdin.close()
            second_process.stdin = None
            first_stdout, first_stderr = first_process.communicate(timeout=30)
            second_stdout, second_stderr = second_process.communicate(timeout=30)

            self.assertEqual(
                first_process.returncode,
                0,
                first_stderr,
            )
            self.assertEqual(
                second_process.returncode,
                0,
                second_stderr,
            )
            self.assertEqual(first_stdout.strip(), "ShadowSuppressed")
            self.assertEqual(second_stdout.strip(), "ShadowSuppressed")
            ledger_data = MODEL.load_cycle_ledger(ledger)
            self.assertEqual(len(ledger_data["entries"]), 2)
            self.assertEqual(
                {
                    entry["response_week"]
                    for entry in ledger_data["entries"]
                },
                {
                    latest_response_week(first_payload).isoformat(),
                    latest_response_week(second_payload).isoformat(),
                },
            )
            self.assertNotIn(
                "RID-",
                ledger.read_text(encoding="utf-8"),
            )

    def test_tampered_cycle_ledger_is_data_blocked_and_not_rewritten(self) -> None:
        payload = make_payload()
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            ledger = root / "state" / "shadow-model-cycle-ledger.json"
            seed_cycle_ledger(ledger, payload)
            tampered = json.loads(ledger.read_text(encoding="utf-8"))
            tampered["entries"][0]["source_sha256"] = "0" * 64
            ledger.write_text(
                json.dumps(tampered, indent=2, sort_keys=True) + "\n",
                encoding="utf-8",
            )
            tampered_bytes = ledger.read_bytes()

            summary = MODEL.run_pipeline(
                payload,
                root / "artifacts",
                {"bootstrap_replicates": 5},
                cycle_ledger_path=ledger,
            )
            self.assertEqual(summary["status"], "DataBlocked")
            self.assertEqual(
                summary["data_blockers"][0]["code"],
                "invalid_cycle_ledger",
            )
            self.assertEqual(ledger.read_bytes(), tampered_bytes)

    def test_cycle_ledger_rejects_unexpected_top_level_and_entry_fields(self) -> None:
        payload = make_payload()
        for mutation in ("top_level", "entry"):
            with self.subTest(mutation=mutation):
                with tempfile.TemporaryDirectory() as temporary_directory:
                    root = Path(temporary_directory)
                    ledger = (
                        root
                        / "state"
                        / "shadow-model-cycle-ledger.json"
                    )
                    seed_cycle_ledger(ledger, payload)
                    ledger_data = json.loads(
                        ledger.read_text(encoding="utf-8")
                    )
                    if mutation == "top_level":
                        ledger_data["unexpected_private_data"] = "blocked"
                    else:
                        ledger_data["entries"][0][
                            "unexpected_private_data"
                        ] = "blocked"
                    ledger.write_text(
                        json.dumps(
                            ledger_data,
                            indent=2,
                            sort_keys=True,
                        )
                        + "\n",
                        encoding="utf-8",
                    )
                    ledger_before = ledger.read_bytes()

                    summary = MODEL.run_pipeline(
                        payload,
                        root / "artifacts",
                        {"bootstrap_replicates": 1},
                        cycle_ledger_path=ledger,
                    )
                    self.assertEqual(summary["status"], "DataBlocked")
                    self.assertEqual(
                        summary["data_blockers"][0]["code"],
                        "invalid_cycle_ledger",
                    )
                    self.assertEqual(ledger.read_bytes(), ledger_before)

    def test_negative_cycles_are_data_blocked(self) -> None:
        payload = make_payload(shadow_cycles=-1)
        with tempfile.TemporaryDirectory() as temporary_directory:
            summary = MODEL.run_pipeline(
                payload,
                Path(temporary_directory),
                {"bootstrap_replicates": 5},
            )
            self.assertEqual(summary["status"], "DataBlocked")
            self.assertEqual(
                summary["data_blockers"][0]["code"],
                "invalid_shadow_cycles",
            )

    def test_nonconverged_fold_models_suppress_all_estimates(self) -> None:
        payload = make_payload()
        original_fit = MODEL.fit_logistic

        def force_nonconvergence(*args, **kwargs):
            beta, _, iterations = original_fit(*args, **kwargs)
            return beta, False, iterations

        MODEL.fit_logistic = force_nonconvergence
        try:
            with tempfile.TemporaryDirectory() as temporary_directory:
                summary, output, _ = run_with_verified_history(
                    payload,
                    Path(temporary_directory),
                    {"bootstrap_replicates": 5},
                )
                self.assertEqual(summary["status"], "ShadowSuppressed")
                self.assertIn(
                    "converged",
                    summary["driver_model"]["reason"],
                )
                estimates = (output / "model_estimates.csv").read_text(
                    encoding="utf-8"
                )
                self.assertEqual(len(estimates.strip().splitlines()), 1)
        finally:
            MODEL.fit_logistic = original_fit

    def test_final_and_bootstrap_nonconvergence_are_exposed(self) -> None:
        payload = make_payload()
        original_fit = MODEL.fit_logistic

        def fail_full_development_fits(x, y, c_value, *args, **kwargs):
            beta, converged, iterations = original_fit(
                x,
                y,
                c_value,
                *args,
                **kwargs,
            )
            if x.shape[0] >= 600:
                converged = False
            return beta, converged, iterations

        MODEL.fit_logistic = fail_full_development_fits
        try:
            with tempfile.TemporaryDirectory() as temporary_directory:
                summary, output, _ = run_with_verified_history(
                    payload,
                    Path(temporary_directory),
                    {"bootstrap_replicates": 5},
                )
                self.assertEqual(summary["status"], "ShadowSuppressed")
                convergence = summary["driver_model"]["convergence"]
                self.assertFalse(convergence["final_model_converged"])
                self.assertEqual(
                    convergence["bootstrap_replicates_requested"],
                    5,
                )
                self.assertEqual(
                    convergence["bootstrap_replicates_completed"],
                    0,
                )
                diagnostics = json.loads(
                    (output / "model_diagnostics.json").read_text(
                        encoding="utf-8"
                    )
                )
                self.assertFalse(
                    diagnostics["models"]["low_overall"]["convergence"][
                        "passed"
                    ]
                )
        finally:
            MODEL.fit_logistic = original_fit

    def test_calibration_nonconvergence_blocks_promotion_criterion(self) -> None:
        payload = make_payload()
        original_fit = MODEL.fit_calibration

        def force_calibration_nonconvergence(*args, **kwargs):
            return None, None, False

        MODEL.fit_calibration = force_calibration_nonconvergence
        try:
            with tempfile.TemporaryDirectory() as temporary_directory:
                summary, _, _ = run_with_verified_history(
                    payload,
                    Path(temporary_directory),
                    {"bootstrap_replicates": 5},
                )
                self.assertEqual(summary["status"], "ShadowReady")
                holdout = summary["driver_model"]["holdout"]
                self.assertFalse(holdout["calibration_converged"])
                self.assertIsNone(holdout["calibration_intercept"])
                self.assertIsNone(holdout["calibration_slope"])
                calibration_check = next(
                    check
                    for check in summary["promotion"]["criteria"]
                    if check["criterion"] == "calibration_model_converged"
                )
                self.assertFalse(calibration_check["met"])
                self.assertFalse(summary["promotion"]["criteria_met"])
        finally:
            MODEL.fit_calibration = original_fit

    def test_cli_and_powershell_wrapper_smoke(self) -> None:
        powershell = shutil.which("powershell") or shutil.which("pwsh")
        if powershell is None:
            self.skipTest("PowerShell is required for the wrapper smoke test.")

        wrapper = REPO_ROOT / "scripts" / "Invoke-GSS-ShadowModel.ps1"
        payload_without_hash = make_payload()
        payload_without_hash.pop("source_sha256")

        def invoke_wrapper(
            input_payload: dict,
            arguments: str,
        ) -> tuple[subprocess.CompletedProcess[str], dict]:
            command = (
                "$result = & "
                f"{quote_powershell_literal(str(wrapper))} {arguments}; "
                "$result | ConvertTo-Json -Compress"
            )
            completed = subprocess.run(
                [
                    powershell,
                    "-NoLogo",
                    "-NoProfile",
                    "-NonInteractive",
                    "-Command",
                    command,
                ],
                input=json.dumps(input_payload, separators=(",", ":")),
                capture_output=True,
                text=True,
                timeout=120,
                check=False,
            )
            self.assertEqual(
                completed.returncode,
                0,
                completed.stderr,
            )
            return completed, json.loads(completed.stdout)

        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            ledger = root / "state" / "shadow-model-cycle-ledger.json"
            controlled_python = quote_powershell_literal(sys.executable)
            ledger_argument = quote_powershell_literal(str(ledger))

            _, hash_result = invoke_wrapper(
                payload_without_hash,
                (
                    "-ComputeSourceHash "
                    f"-PythonPath {controlled_python} "
                    f"-CycleLedgerPath {ledger_argument}"
                ),
            )
            self.assertEqual(
                hash_result["Status"],
                "HashComputed",
                hash_result.get("TechnicalError"),
            )
            self.assertRegex(hash_result["SourceSha256"], r"^[0-9a-f]{64}$")
            self.assertFalse(ledger.exists())

            payload = copy.deepcopy(payload_without_hash)
            payload["source_sha256"] = hash_result["SourceSha256"]
            output = root / "artifacts"
            _, run_result = invoke_wrapper(
                payload,
                (
                    f"-OutputDirectory {quote_powershell_literal(str(output))} "
                    f"-PythonPath {controlled_python} "
                    f"-CycleLedgerPath {ledger_argument}"
                ),
            )
            self.assertEqual(
                run_result["Status"],
                "ShadowSuppressed",
                run_result.get("TechnicalError"),
            )
            self.assertEqual(
                sorted(path.name for path in output.iterdir()),
                sorted(MODEL.REQUIRED_ARTIFACTS),
            )
            self.assertEqual(
                Path(run_result["CycleLedgerPath"]).resolve(),
                ledger.resolve(),
            )
            self.assertTrue(ledger.exists())

            forbidden_output = REPO_ROOT / "wrapper-smoke-forbidden"
            _, guard_result = invoke_wrapper(
                payload,
                (
                    "-OutputDirectory "
                    f"{quote_powershell_literal(str(forbidden_output))} "
                    f"-PythonPath {controlled_python} "
                    f"-CycleLedgerPath {ledger_argument}"
                ),
            )
            self.assertEqual(guard_result["Status"], "TechnicalError")
            self.assertIn(
                "outside the program repository",
                guard_result["TechnicalError"],
            )
            self.assertFalse(forbidden_output.exists())

            missing_python = root / "missing-python.exe"
            missing_ledger = root / "missing-state" / "ledger.json"
            _, missing_runtime_result = invoke_wrapper(
                payload,
                (
                    "-OutputDirectory "
                    f"{quote_powershell_literal(str(root / 'missing-output'))} "
                    "-PythonPath "
                    f"{quote_powershell_literal(str(missing_python))} "
                    "-CycleLedgerPath "
                    f"{quote_powershell_literal(str(missing_ledger))}"
                ),
            )
            self.assertEqual(
                missing_runtime_result["Status"],
                "TechnicalError",
            )
            self.assertFalse(missing_ledger.exists())


if __name__ == "__main__":
    unittest.main()
