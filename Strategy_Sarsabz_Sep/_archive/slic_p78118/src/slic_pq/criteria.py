"""Criteria encoded from the SLIC prequalification document (ref P78118, 18 Aug 2026).

Source: data/pq-bidding-document (2).pdf -- section "Eligibility & Qualification
Criteria" (p.21-22) and "Evaluation Criteria" (p.22-23). Every item carries the
page it was read from so a reviewer can trace it back.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Literal

Bucket = Literal["compliance", "capability", "pitch"]


@dataclass(frozen=True)
class Tender:
    reference: str = "P78118"
    title: str = "Prequalification for Procurement of Advertising Agencies"
    agency: str = "State Life Insurance Corporation of Pakistan (SLIC)"
    division: str = (
        "Central Procurement Division, State Life Building No 09, "
        "Dr Ziauddin Ahmed Road, Karachi"
    )
    contact_email: str = "dmgs@statelife.com.pk"
    contact_phone: str = "+92-300-228-8647"
    portal: str = "https://epads.gov.pk/opportunities/federal/procurements/78118"
    vendor_registration: str = "https://vendors.epads.gov.pk/"
    method: str = "Single Stage - One Envelope, National"
    selection: str = "Quality Based Selection (QBS)"
    engagement_term: str = "2026-28"
    issued: str = "2026-08-18"
    clarification_deadline: str = "2026-08-31"
    submission_deadline: str = "2026-09-03 11:00"
    opening: str = "2026-09-03 11:30"
    total_marks: int = 100
    passing_marks: int = 50


@dataclass(frozen=True)
class EligibilityItem:
    """A pass/fail document the application must carry (PQ doc p.21-22)."""

    key: str
    requirement: str
    page: int = 21


@dataclass(frozen=True)
class ScoredItem:
    """A marked criterion in the technical evaluation (PQ doc p.22-23)."""

    key: str
    question: str
    marks: int
    bucket: Bucket
    evidence: str = ""
    page: int = 22


ELIGIBILITY: tuple[EligibilityItem, ...] = (
    EligibilityItem(
        "profile",
        "Agency/firm profile: name, registered address, telephone, fax, e-mail of head "
        "office and branch offices, and year of establishment.",
    ),
    EligibilityItem(
        "apns_pba_aap",
        "Registration certificates with All Pakistan Newspapers Society (APNS), Pakistan "
        "Broadcasting Association (PBA) and Advertising Association of Pakistan (AAP).",
    ),
    EligibilityItem(
        "pid",
        "Enlistment certificate / letter of Press Information Department (PID).",
    ),
    EligibilityItem(
        "tech_staff",
        "Particulars of permanent technical staff -- qualification, experience and "
        "available facilities.",
    ),
    EligibilityItem(
        "not_blacklisted",
        "Certificate that the agency has not been blacklisted/suspended by APNS, PBA etc.",
    ),
    EligibilityItem("tax_cert", "Income Tax / GST payment certificate."),
    EligibilityItem("bank_cert", "Bank certificate of financial stability."),
    EligibilityItem(
        "foreign_affiliation", "Foreign affiliation / associates, if any."
    ),
    EligibilityItem(
        "client_list",
        "List of clients and details of services offered to them during the last five "
        "(5) years.",
        page=22,
    ),
    EligibilityItem(
        "slic_dues_clearance",
        "Certificate that all bills/invoices to newspapers have been cleared/paid, where "
        "engaged or released by State Life.",
        page=22,
    ),
)

REGISTRATIONS: tuple[str, ...] = ("FBR (NTN)", "FBR (GSTN)")

BIDDER_TYPES: tuple[str, ...] = (
    "Individual / Individual Consultant",
    "Partnership Firm",
    "Company (Private Limited)",
    "Company (Public Limited)",
)

SCORED: tuple[ScoredItem, ...] = (
    ScoredItem(
        "secp",
        'Is the agency incorporated as "Private Limited" with SECP?',
        2,
        "compliance",
        "Attested SECP Certificate of Incorporation, Memorandum and Articles of Association.",
    ),
    ScoredItem(
        "fbr",
        "Is the agency registered with FBR and/or provincial revenue board(s) and an "
        "active taxpayer?",
        2,
        "compliance",
        "Attested NTN plus tax returns for 2023, 2024, 2025 and FBR certificate.",
    ),
    ScoredItem(
        "accreditation",
        "Is the agency accredited with APNS and PBA, and enlisted by PID?",
        3,
        "compliance",
        "APNS and PBA accreditation certificates, PID enlistment letter.",
    ),
    ScoredItem(
        "clean_record",
        "Does the agency confirm it has never been involved in criminal/unlawful "
        "activity nor been blacklisted by any entity/APNS/PBA?",
        3,
        "compliance",
        "Declaration on Rs. 100 stamp paper per Annexure II.",
    ),
    ScoredItem(
        "campaign_experience",
        "Does the agency have the last three (03) years experience in launching "
        "advertising campaigns?",
        3,
        "capability",
        "Published campaign material: print ads, TVC/DVC/URLs, invoices, transmission "
        "certificates.",
        page=23,
    ),
    ScoredItem(
        "active_clients",
        "Does the agency have a minimum of three (03) active clients at time of submission?",
        2,
        "capability",
        "Contracts, current POs or client confirmation letters.",
        page=23,
    ),
    ScoredItem(
        "revenue",
        "Does the agency have minimum aggregate revenue of PKR 100m-500m, or above "
        "500m, for the last 03 years?",
        4,
        "capability",
        "Audited accounts / statutory financials for the last three years.",
        page=23,
    ),
    ScoredItem(
        "expertise",
        "Does the agency possess expertise in art direction, creative direction, "
        "copywriting and account/client management?",
        2,
        "capability",
        "CVs detailing technical capability, qualification and experience, per Annexure III.",
        page=23,
    ),
    ScoredItem(
        "offices",
        "Does the agency have an operational office in Karachi, Lahore and Islamabad?",
        2,
        "capability",
        "Lease deeds, utility bills or letterheads evidencing all three offices.",
        page=23,
    ),
    ScoredItem(
        "headcount",
        "Does the agency have permanent employees from 20 to 50, or above?",
        4,
        "capability",
        "Payroll summary / EOBI or SESSI records.",
        page=23,
    ),
    ScoredItem(
        "tech_staff_particulars",
        "Particulars of the permanent technical staff -- qualification, experience and "
        "available facilities.",
        3,
        "capability",
        "Annexure III staff table with enclosed CVs and employment status.",
        page=23,
    ),
    ScoredItem(
        "campaign_concept",
        "Campaign concept.",
        15,
        "pitch",
        "Written concept for the SLIC campaign, with rationale.",
        page=23,
    ),
    ScoredItem(
        "competitive_analysis",
        "Competitive analysis.",
        15,
        "pitch",
        "Category and competitor review for life insurance in Pakistan.",
        page=23,
    ),
    ScoredItem(
        "communication_strategy",
        "Communication strategy.",
        15,
        "pitch",
        "Audience, positioning, message architecture and channel role.",
        page=23,
    ),
    ScoredItem(
        "creative_artworks",
        "Creative artworks.",
        15,
        "pitch",
        "Executions across the proposed channels.",
        page=23,
    ),
    ScoredItem(
        "media_mix",
        "Proposed media mix.",
        10,
        "pitch",
        "Channel split, weights and indicative budget logic.",
        page=23,
    ),
)

BUCKET_LABELS: dict[str, str] = {
    "compliance": "Corporate compliance (documents)",
    "capability": "Agency capability & scale",
    "pitch": "Campaign submission (subjective)",
}


def total_marks() -> int:
    return sum(item.marks for item in SCORED)


def marks_by_bucket() -> dict[str, int]:
    out: dict[str, int] = {}
    for item in SCORED:
        out[item.bucket] = out.get(item.bucket, 0) + item.marks
    return out


def scored_by_key() -> dict[str, ScoredItem]:
    return {item.key: item for item in SCORED}


def as_dict() -> dict:
    return {
        "tender": asdict(Tender()),
        "bidder_types": list(BIDDER_TYPES),
        "registrations": list(REGISTRATIONS),
        "eligibility": [asdict(i) for i in ELIGIBILITY],
        "scored": [asdict(i) for i in SCORED],
        "marks_by_bucket": marks_by_bucket(),
        "total_marks": total_marks(),
    }
