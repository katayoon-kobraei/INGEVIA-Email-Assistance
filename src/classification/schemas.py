from enum import Enum
from pydantic import BaseModel


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
    is_relevant: bool  # false = social media / marketing / automated noise -- not real client correspondence


class PendingResult(BaseModel):
    needs_response: bool  # true = this email is waiting on a written reply from the firm


class ProjectMatchResult(BaseModel):
    matched_existing: bool
    project_folder_name: str
    contact_label: str   # e.g. "E.BOSCH" or "DAVID REYERO (REYQUEDA)"
    topic_label: str
    mentions_specific_address: bool  # true if THIS email names a specific site/address/location


class AddressMatchResult(BaseModel):
    matched_existing: bool
    address_folder_name: str   # bare name, no code prefix -- e.g. "CAMÍ DE FAITANAR 2 - PICAÑA"