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


class ProjectMatchResult(BaseModel):
    matched_existing: bool
    project_folder_name: str
    contact_label: str   # e.g. "E.BOSCH" or "DAVID REYERO (REYQUEDA)"
    topic_label: str