from enum import Enum
from pydantic import BaseModel, Field


class Priority(str, Enum):
    HIGH = "High"
    MEDIUM = "Medium"
    LOW = "Low"


class Department(str, Enum):
    HR = "HR"
    FINANCE = "Finance"
    PROCUREMENT = "Procurement"
    PROJECTS = "Projects"
    SALES = "Sales"
    LEGAL = "Legal"
    GENERAL = "General"


class PriorityResult(BaseModel):
    priority: Priority
    reasoning: str


class DepartmentResult(BaseModel):
    department: Department
    reasoning: str


class RelevanceResult(BaseModel):
    is_relevant: bool  # unused now -- kept for the orphaned relevance_agent.py


class PendingResult(BaseModel):
    needs_response: bool  # true = this email is waiting on a written reply from the firm
    # kept for compatibility -- no longer called from pipeline.py, superseded by PostFilingResult


class PostFilingResult(BaseModel):
    """Merges the old separate pending-check and priority-score calls
    into one -- both are cheap, informational-only checks that need
    nothing but the email itself, so there's no reason to pay for the
    email body twice."""
    needs_response: bool
    priority: int = Field(ge=1, le=5)


class ProjectMatchResult(BaseModel):
    is_relevant: bool  # true = real project correspondence or official bank/government mail worth filing; false = everything else -- left untouched, staff handle it manually
    matched_existing: bool
    project_folder_name: str
    contact_label: str   # e.g. "E.BOSCH" or "DAVID REYERO (REYQUEDA)"
    topic_label: str
    mentions_specific_address: bool  # true if THIS email names a specific site/address/location
    summary: str  # one-line, plain-Spanish (not ALL-CAPS) summary for the human-readable tracking report -- only meaningful when is_relevant=True


class AddressMatchResult(BaseModel):
    matched_existing: bool
    address_folder_name: str   # bare name, no code prefix -- e.g. "CAMÍ DE FAITANAR 2 - PICAÑA"


class BillingMatchResult(BaseModel):
    is_billing_related: bool  # true = factura, proforma, oferta, presupuesto, licitación, or contratación del Estado


class PriorityScoreResult(BaseModel):
    priority: int = Field(ge=1, le=5)  # 1 = low urgency, 5 = urgent


class PlenergyAddressMatchResult(BaseModel):
    matched_existing: bool
    matched_folder: str      # "26-004 DO PLENERGY" or "26-003 PLENERGY" -- only meaningful when matched_existing=True
    address_folder_name: str # exact existing name if matched; otherwise a NEW bare site name proposal
    contact_name: str        # ALL CAPS contact name from signature/greeting -- only meaningful when matched_existing=False; "" otherwise
    summary: str             # one-line, plain-Spanish summary for the human-readable tracking report