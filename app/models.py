"""Pydantic モデル: Web UI フォームのバリデーション."""

from pydantic import BaseModel


class AnalysisConfig(BaseModel):
    company_domains: list[str] = []
    cc_key_person_threshold: float = 0.30
    min_edge_weight: int = 1
    hub_degree_weight: float = 0.5
    hub_betweenness_weight: float = 0.5
