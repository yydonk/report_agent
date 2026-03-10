from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any, Optional
import json


@dataclass
class OCRBlock:
    text: str
    bbox: List[int] = field(default_factory=list)
    confidence: Optional[float] = None
    page: Optional[int] = None
    block_type: Optional[str] = None


@dataclass
class OCRResult:
    blocks: List[OCRBlock] = field(default_factory=list)
    raw_text: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {"blocks": [asdict(b) for b in self.blocks], "raw_text": self.raw_text}


@dataclass
class TableData:
    name: Optional[str] = None
    columns: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    units: Optional[Dict[str, str]] = None


@dataclass
class StepItem:
    step_id: Optional[str] = None
    description: str = ""
    parameters: Optional[Dict[str, Any]] = None
    observation: Optional[str] = None
    notes: Optional[str] = None


@dataclass
class StructuredExperiment:
    title: str
    objective: str
    theory: str
    apparatus: List[str]
    steps: List[StepItem]
    data: Dict[str, Any]
    analysis: str

    def to_json(self) -> str:
        return json.dumps(
            {
                "title": self.title,
                "objective": self.objective,
                "theory": self.theory,
                "apparatus": self.apparatus,
                "steps": [asdict(s) for s in self.steps],
                "data": self.data,
                "analysis": self.analysis,
            },
            ensure_ascii=False,
            indent=2,
        )


def empty_structured() -> StructuredExperiment:
    return StructuredExperiment(
        title="",
        objective="",
        theory="",
        apparatus=[],
        steps=[],
        data={"tables": [], "observations": [], "raw_text": ""},
        analysis="",
    )
