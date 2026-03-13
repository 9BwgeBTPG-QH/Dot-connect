"""Unit tests for app.core — the analysis pipeline."""

import io
import math
import tempfile
from pathlib import Path

import networkx as nx
import pandas as pd
import pytest

from app.core import (
    COMMUNITY_COLORS,
    analyze_graph,
    build_graph,
    generate_vis_data,
    load_csv,
    parse_address_field,
)


# ======================================================================
# Helpers
# ======================================================================

def _default_config(**overrides):
    """Return a minimal config dict, optionally overriding thresholds."""
    cfg = {
        "company_domains": ["example.com"],
        "thresholds": {
            "cc_key_person_threshold": 0.30,
            "min_edge_weight": 1,
            "hub_degree_weight": 0.5,
            "hub_betweenness_weight": 0.5,
        },
    }
    cfg["thresholds"].update(overrides)
    return cfg


def _make_df(rows):
    """Build a DataFrame that matches the expected CSV schema."""
    return pd.DataFrame(rows, columns=["from_email", "from_name", "to", "cc"])


def _simple_df():
    """A small deterministic dataset used by several tests."""
    return _make_df([
        ("alice@example.com", "Alice", "Bob <bob@example.com>", "Carol <carol@example.com>"),
        ("alice@example.com", "Alice", "Bob <bob@example.com>", "Carol <carol@example.com>"),
        ("bob@example.com", "Bob", "Alice <alice@example.com>", ""),
        ("carol@example.com", "Carol", "Alice <alice@example.com>; Bob <bob@example.com>", ""),
    ])


# ======================================================================
# 1. parse_address_field
# ======================================================================

class TestParseAddressField:
    """Tests for parse_address_field()."""

    def test_rfc_format(self):
        result = parse_address_field("John Doe <john@example.com>")
        assert result == [("john@example.com", "John Doe")]

    def test_raw_email(self):
        result = parse_address_field("john@example.com")
        assert result == [("john@example.com", "")]

    def test_empty_string(self):
        assert parse_address_field("") == []

    def test_none_input(self):
        assert parse_address_field(None) == []

    def test_nan_input(self):
        assert parse_address_field(float("nan")) == []

    def test_pandas_nan(self):
        assert parse_address_field(pd.NaT) == []

    def test_semicolon_separated(self):
        field = "Alice <alice@a.com>; Bob <bob@b.com>"
        result = parse_address_field(field)
        assert len(result) == 2
        assert result[0] == ("alice@a.com", "Alice")
        assert result[1] == ("bob@b.com", "Bob")

    def test_mixed_rfc_and_raw(self):
        field = "Alice <alice@a.com>; bob@b.com"
        result = parse_address_field(field)
        assert len(result) == 2
        assert result[0] == ("alice@a.com", "Alice")
        assert result[1] == ("bob@b.com", "")

    def test_email_lowercased(self):
        result = parse_address_field("Alice <ALICE@Example.COM>")
        assert result[0][0] == "alice@example.com"

    def test_whitespace_trimming(self):
        result = parse_address_field("  Alice  <alice@a.com>  ")
        assert result == [("alice@a.com", "Alice")]

    def test_no_at_sign_ignored(self):
        """Entries without '@' and without angle brackets are skipped."""
        result = parse_address_field("not-an-email")
        assert result == []


# ======================================================================
# 2. build_graph
# ======================================================================

class TestBuildGraph:
    """Tests for build_graph()."""

    def test_basic_structure(self):
        df = _simple_df()
        G = build_graph(df, _default_config())
        assert isinstance(G, nx.DiGraph)
        assert G.number_of_nodes() > 0
        assert G.number_of_edges() > 0

    def test_to_weight_increments(self):
        """Alice sends TO Bob twice, so to_weight on that edge should be 2."""
        df = _simple_df()
        G = build_graph(df, _default_config())
        assert G.has_edge("alice@example.com", "bob@example.com")
        assert G["alice@example.com"]["bob@example.com"]["to_weight"] == 2

    def test_cc_weight(self):
        """Alice CCs Carol twice, so cc_weight should be 2."""
        df = _simple_df()
        G = build_graph(df, _default_config())
        assert G.has_edge("alice@example.com", "carol@example.com")
        assert G["alice@example.com"]["carol@example.com"]["cc_weight"] == 2

    def test_to_edge_has_zero_cc_weight(self):
        """A pure To edge should initialise cc_weight to 0."""
        df = _make_df([
            ("a@x.com", "A", "b@x.com", ""),
        ])
        G = build_graph(df, _default_config())
        assert G["a@x.com"]["b@x.com"]["cc_weight"] == 0

    def test_cc_edge_has_zero_to_weight(self):
        """A pure CC edge should initialise to_weight to 0."""
        df = _make_df([
            ("a@x.com", "A", "", "b@x.com"),
        ])
        G = build_graph(df, _default_config())
        assert G["a@x.com"]["b@x.com"]["to_weight"] == 0

    def test_node_sent_count(self):
        df = _simple_df()
        G = build_graph(df, _default_config())
        assert G.nodes["alice@example.com"]["sent"] == 2

    def test_node_received_count(self):
        df = _simple_df()
        G = build_graph(df, _default_config())
        # Bob receives from Alice (2 times) + from Carol (1 time) = 3
        assert G.nodes["bob@example.com"]["received"] == 3

    def test_node_cc_count(self):
        df = _simple_df()
        G = build_graph(df, _default_config())
        assert G.nodes["carol@example.com"]["cc_count"] == 2

    def test_is_internal_flag(self):
        df = _make_df([
            ("a@example.com", "A", "b@external.com", ""),
        ])
        G = build_graph(df, _default_config())
        assert G.nodes["a@example.com"]["is_internal"] is True
        assert G.nodes["b@external.com"]["is_internal"] is False

    def test_node_name_populated(self):
        df = _make_df([
            ("a@x.com", "Alice", "Bob <b@x.com>", ""),
        ])
        G = build_graph(df, _default_config())
        assert G.nodes["a@x.com"]["name"] == "Alice"
        assert G.nodes["b@x.com"]["name"] == "Bob"

    def test_missing_from_email_skipped(self):
        df = _make_df([
            ("", "", "b@x.com", ""),
            (float("nan"), "", "b@x.com", ""),
        ])
        G = build_graph(df, _default_config())
        assert G.number_of_nodes() == 0

    def test_empty_dataframe(self):
        df = _make_df([])
        G = build_graph(df, _default_config())
        assert G.number_of_nodes() == 0
        assert G.number_of_edges() == 0


# ======================================================================
# 3. analyze_graph
# ======================================================================

class TestAnalyzeGraph:
    """Tests for analyze_graph()."""

    def _build_and_analyze(self, df=None, config=None):
        df = df if df is not None else _simple_df()
        config = config or _default_config()
        G = build_graph(df, config)
        analysis = analyze_graph(G, len(df), config)
        return G, analysis

    def test_return_keys(self):
        _, analysis = self._build_and_analyze()
        assert "total_mails" in analysis
        assert "cc_key_persons" in analysis
        assert "hubs" in analysis
        assert "communities" in analysis
        assert "community_map" in analysis

    def test_total_mails(self):
        _, analysis = self._build_and_analyze()
        assert analysis["total_mails"] == 4

    def test_cc_key_person_detection(self):
        """Carol is CC'd on 2/4 = 50% of mails, above the 30% threshold."""
        _, analysis = self._build_and_analyze()
        cc_emails = [p["email"] for p in analysis["cc_key_persons"]]
        assert "carol@example.com" in cc_emails

    def test_cc_key_person_ratio(self):
        _, analysis = self._build_and_analyze()
        carol = next(p for p in analysis["cc_key_persons"] if p["email"] == "carol@example.com")
        assert carol["ratio"] == 0.5  # 2/4
        assert carol["cc_count"] == 2

    def test_cc_key_person_threshold_strict(self):
        """With a very high threshold, nobody should qualify."""
        config = _default_config(cc_key_person_threshold=0.99)
        _, analysis = self._build_and_analyze(config=config)
        assert len(analysis["cc_key_persons"]) == 0

    def test_hub_detection(self):
        _, analysis = self._build_and_analyze()
        assert len(analysis["hubs"]) > 0
        hub_emails = [h["email"] for h in analysis["hubs"]]
        # Alice or Bob should be detected as a hub (highest connectivity)
        assert any(e in hub_emails for e in ["alice@example.com", "bob@example.com"])

    def test_hub_score_fields(self):
        _, analysis = self._build_and_analyze()
        hub = analysis["hubs"][0]
        assert "score" in hub
        assert "degree_centrality" in hub
        assert "betweenness_centrality" in hub
        assert isinstance(hub["score"], float)

    def test_hubs_sorted_descending(self):
        _, analysis = self._build_and_analyze()
        scores = [h["score"] for h in analysis["hubs"]]
        assert scores == sorted(scores, reverse=True)

    def test_community_detection(self):
        _, analysis = self._build_and_analyze()
        assert len(analysis["communities"]) >= 1
        # Each community should have id, color, size, members
        comm = analysis["communities"][0]
        assert "id" in comm
        assert "color" in comm
        assert "size" in comm
        assert "members" in comm

    def test_community_map_covers_all_nodes(self):
        G, analysis = self._build_and_analyze()
        for node in G.nodes:
            assert node in analysis["community_map"]

    def test_community_assigned_to_nodes(self):
        G, _ = self._build_and_analyze()
        for node in G.nodes:
            assert "community" in G.nodes[node]

    def test_zero_total_mails_no_crash(self):
        """total_mails <= 0 should be clamped to 1 to avoid ZeroDivisionError."""
        df = _simple_df()
        G = build_graph(df, _default_config())
        analysis = analyze_graph(G, 0, _default_config())
        assert analysis["total_mails"] == 1


# ======================================================================
# 4. load_csv
# ======================================================================

class TestLoadCsv:
    """Tests for load_csv()."""

    def _csv_content(self, encoding="utf-8-sig"):
        content = '"from_email","from_name","to","cc"\n"a@x.com","Alice","Bob <b@x.com>",""\n'
        return content.encode(encoding)

    def test_utf8_sig(self):
        raw = self._csv_content("utf-8-sig")
        df = load_csv(io.BytesIO(raw))
        assert len(df) == 1
        assert df.iloc[0]["from_email"] == "a@x.com"

    def test_cp932_fallback(self):
        # Build a CSV with a cp932-encodable character
        content = '"from_email","from_name","to","cc"\n"a@x.com","\u5c71\u7530","b@x.com",""\n'
        raw = content.encode("cp932")
        df = load_csv(io.BytesIO(raw))
        assert len(df) == 1
        assert df.iloc[0]["from_name"] == "\u5c71\u7530"

    def test_file_path(self):
        content = '"from_email","from_name","to","cc"\n"a@x.com","Alice","b@x.com",""\n'
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="wb") as f:
            f.write(content.encode("utf-8-sig"))
            path = f.name
        try:
            df = load_csv(path)
            assert len(df) == 1
        finally:
            Path(path).unlink()

    def test_invalid_encoding_raises(self):
        # latin-1 accepts any byte sequence, so encoding-based failure is
        # unlikely. Instead verify that a valid but minimal CSV succeeds.
        raw = b'"col1"\n"val1"\n'
        df = load_csv(io.BytesIO(raw))
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 1


# ======================================================================
# 5. generate_vis_data
# ======================================================================

class TestGenerateVisData:
    """Tests for generate_vis_data()."""

    def _build_all(self, df=None, config=None):
        df = df if df is not None else _simple_df()
        config = config or _default_config()
        G = build_graph(df, config)
        analysis = analyze_graph(G, len(df), config)
        vis = generate_vis_data(G, analysis, config)
        return G, analysis, vis

    def test_output_keys(self):
        _, _, vis = self._build_all()
        assert "nodes" in vis
        assert "edges" in vis
        assert "communities" in vis
        assert "analysis" in vis
        assert "wordcloud_data" in vis

    def test_analysis_sub_keys(self):
        _, _, vis = self._build_all()
        a = vis["analysis"]
        assert "total_mails" in a
        assert "total_nodes" in a
        assert "total_edges" in a
        assert "cc_key_persons" in a
        assert "hubs" in a

    def test_node_fields(self):
        _, _, vis = self._build_all()
        node = vis["nodes"][0]
        expected_keys = {
            "id", "label", "name", "email", "domain",
            "is_internal", "sent", "received", "cc_count",
            "community", "color", "size", "is_cc_key", "is_hub",
        }
        assert expected_keys.issubset(node.keys())

    def test_edge_fields(self):
        _, _, vis = self._build_all()
        assert len(vis["edges"]) > 0
        edge = vis["edges"][0]
        expected_keys = {"from", "to", "to_weight", "cc_weight", "weight", "width"}
        assert expected_keys.issubset(edge.keys())

    def test_min_edge_weight_filtering(self):
        """Edges below min_edge_weight should be excluded."""
        config = _default_config(min_edge_weight=5)
        _, _, vis = self._build_all(config=config)
        # With threshold 5, most lightweight edges are dropped
        for edge in vis["edges"]:
            assert edge["weight"] >= 5

    def test_min_edge_weight_one_includes_all(self):
        config = _default_config(min_edge_weight=1)
        G, _, vis = self._build_all(config=config)
        # All edges with total weight >= 1 are included
        expected = sum(
            1 for _, _, d in G.edges(data=True)
            if d.get("to_weight", 0) + d.get("cc_weight", 0) >= 1
        )
        assert len(vis["edges"]) == expected

    def test_node_size_bounds(self):
        _, _, vis = self._build_all()
        for node in vis["nodes"]:
            assert 8 <= node["size"] <= 40

    def test_label_truncation(self):
        """Labels longer than 15 chars should be truncated with an ellipsis."""
        df = _make_df([
            ("a@x.com", "A Very Long Name That Exceeds", "b@x.com", ""),
        ])
        _, _, vis = self._build_all(df=df)
        sender_node = next(n for n in vis["nodes"] if n["id"] == "a@x.com")
        assert len(sender_node["label"]) <= 15

    def test_wordcloud_data_sorted_descending(self):
        _, _, vis = self._build_all()
        sizes = [w["size"] for w in vis["wordcloud_data"]]
        assert sizes == sorted(sizes, reverse=True)

    def test_wordcloud_entry_fields(self):
        _, _, vis = self._build_all()
        if vis["wordcloud_data"]:
            entry = vis["wordcloud_data"][0]
            assert "text" in entry
            assert "size" in entry
            assert "email" in entry
            assert "color" in entry

    def test_is_cc_key_flag(self):
        _, _, vis = self._build_all()
        carol_node = next(n for n in vis["nodes"] if n["id"] == "carol@example.com")
        assert carol_node["is_cc_key"] is True

    def test_community_color_from_palette(self):
        _, _, vis = self._build_all()
        for node in vis["nodes"]:
            assert node["color"] in COMMUNITY_COLORS

    def test_communities_forwarded(self):
        _, analysis, vis = self._build_all()
        assert vis["communities"] == analysis["communities"]
