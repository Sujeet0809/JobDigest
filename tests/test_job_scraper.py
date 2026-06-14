"""
Test suite for the Product Leadership Job Digest scraper.

These tests cover the pure, network-free logic: configuration loading,
title filtering, relevancy scoring, country/work-mode detection,
deduplication, and JSON-LD parsing. They are written test-first (TDD)
and must all pass before the scraper is wired into CI.
"""

import os
import sys

from bs4 import BeautifulSoup

# Make the scraper importable when tests run from the repo root.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import job_scraper as js  # noqa: E402


# ─────────────────────────────────────────────
# load_config — secrets come from the environment, never hard-coded
# ─────────────────────────────────────────────

def test_load_config_reads_password_from_env():
    cfg = js.load_config({
        "GMAIL_APP_PASS": "abcd efgh ijkl mnop",
        "SENDER_EMAIL": "me@example.com",
        "RECEIVER_EMAIL": "you@example.com",
    })
    assert cfg["app_pass"] == "abcd efgh ijkl mnop"
    assert cfg["sender"] == "me@example.com"
    assert cfg["receiver"] == "you@example.com"


def test_load_config_receiver_defaults_to_sender():
    cfg = js.load_config({
        "GMAIL_APP_PASS": "secret",
        "SENDER_EMAIL": "solo@example.com",
    })
    assert cfg["receiver"] == "solo@example.com"


def test_load_config_missing_password_is_empty_not_placeholder():
    cfg = js.load_config({})
    # No password configured → empty string so main() can skip cleanly.
    assert cfg["app_pass"] == ""
    # Must never fall back to a committed literal password.
    assert "YOUR_GMAIL_APP_PASSWORD" not in cfg.values()


def test_no_hardcoded_password_in_source():
    src = open(js.__file__, encoding="utf-8").read()
    assert "YOUR_GMAIL_APP_PASSWORD" not in src


# ─────────────────────────────────────────────
# is_leadership — title filtering
# ─────────────────────────────────────────────

def test_is_leadership_accepts_leadership_titles():
    for title in [
        "VP of Product",
        "Director of Product",
        "Head of Product",
        "Chief Product Officer",
        "Group Product Manager",
        "Principal Product Manager",
        "Senior Product Manager",
        "Product Manager - Growth",
    ]:
        assert js.is_leadership(title), title


def test_is_leadership_rejects_non_product_titles():
    for title in [
        "Software Engineer",
        "Data Analyst",
        "Project Coordinator",
        "Marketing Manager",
        "",
    ]:
        assert not js.is_leadership(title), title


def test_is_leadership_is_case_insensitive():
    assert js.is_leadership("senior PRODUCT manager")


# ─────────────────────────────────────────────
# score_job — relevancy 1..10
# ─────────────────────────────────────────────

def test_score_is_clamped_between_1_and_10():
    for job in [
        {"title": "VP of Product", "company": "HealthCo", "location": "Bangalore"},
        {"title": "Intern", "company": "Gaming Studio", "location": "USA"},
        {"title": "", "company": "", "location": ""},
    ]:
        s = js.score_job(job)
        assert 1 <= s <= 10


def test_strong_industry_and_india_outscore_bad_industry_and_us():
    strong = js.score_job({
        "title": "Director of Product",
        "company": "MedTech Health Insurance",
        "location": "Bengaluru, India",
    })
    weak = js.score_job({
        "title": "Associate Product Manager",
        "company": "Online Gaming Studio",
        "location": "San Francisco, USA",
    })
    assert strong > weak


def test_seniority_boosts_score():
    senior = js.score_job({"title": "VP of Product", "company": "X", "location": "Remote"})
    junior = js.score_job({"title": "Junior Product Manager", "company": "X", "location": "Remote"})
    assert senior > junior


# ─────────────────────────────────────────────
# detect_country_flag
# ─────────────────────────────────────────────

def test_detect_country_flag():
    assert js.detect_country_flag("Bengaluru, India") == "🇮🇳"
    assert js.detect_country_flag("New York, USA") == "🇺🇸"
    assert js.detect_country_flag("London, UK") == "🇬🇧"
    assert js.detect_country_flag("Remote") == "🌍"
    assert js.detect_country_flag("") == "🌐"


# ─────────────────────────────────────────────
# detect_work_mode
# ─────────────────────────────────────────────

def test_detect_work_mode():
    assert js.detect_work_mode("Senior PM", "Remote - India")[1] == "Remote"
    assert js.detect_work_mode("Senior PM (Hybrid)", "Mumbai")[1] == "Hybrid"
    assert js.detect_work_mode("Senior PM", "On-site Pune")[1] == "On-site"
    # Unknown defaults to Hybrid.
    assert js.detect_work_mode("Senior PM", "Mumbai")[1] == "Hybrid"


# ─────────────────────────────────────────────
# deduplicate
# ─────────────────────────────────────────────

def test_deduplicate_removes_same_title_company():
    jobs = [
        {"title": "VP Product", "company": "Acme"},
        {"title": "vp product", "company": "ACME"},   # same, different case
        {"title": "VP Product", "company": "Globex"}, # different company
    ]
    out = js.deduplicate(jobs)
    assert len(out) == 2


def test_deduplicate_preserves_first_occurrence_order():
    jobs = [
        {"title": "A", "company": "1"},
        {"title": "B", "company": "2"},
        {"title": "A", "company": "1"},
    ]
    out = js.deduplicate(jobs)
    assert [j["title"] for j in out] == ["A", "B"]


# ─────────────────────────────────────────────
# parse_jsonld_jobs
# ─────────────────────────────────────────────

def test_parse_jsonld_extracts_leadership_jobs_only():
    html = """
    <html><head>
    <script type="application/ld+json">
    {"@type":"JobPosting","title":"Director of Product",
     "hiringOrganization":{"name":"Acme"},
     "jobLocation":{"address":{"addressLocality":"Bengaluru"}},
     "url":"https://acme.test/job/1","datePosted":"2026-06-10"}
    </script>
    <script type="application/ld+json">
    {"@type":"JobPosting","title":"Backend Engineer",
     "hiringOrganization":{"name":"Acme"}}
    </script>
    </head><body></body></html>
    """
    soup = BeautifulSoup(html, "html.parser")
    jobs = js.parse_jsonld_jobs(soup, "TestSource", "https://fallback.test")
    assert len(jobs) == 1
    job = jobs[0]
    assert job["title"] == "Director of Product"
    assert job["company"] == "Acme"
    assert job["location"] == "Bengaluru"
    assert job["source"] == "TestSource"
    assert job["date"] == "2026-06-10"


def test_parse_jsonld_handles_malformed_json_gracefully():
    html = '<script type="application/ld+json">{ not valid json }</script>'
    soup = BeautifulSoup(html, "html.parser")
    assert js.parse_jsonld_jobs(soup, "S", "https://f.test") == []
