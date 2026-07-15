"""Packaged entry point — always pauses so a double-click console can be read."""

import sys

import networth._packaged as P


def _capture(captured):
    def fake(argv):
        captured["argv"] = list(argv)
        return 0
    return fake


def test_run_appends_pause(monkeypatch):
    captured = {}
    monkeypatch.setattr(P, "main", _capture(captured))
    monkeypatch.setattr(sys, "argv", ["Update Portfolio"])
    assert P.run() == 0
    assert captured["argv"] == ["--pause"]


def test_run_preserves_args_without_doubling_pause(monkeypatch):
    captured = {}
    monkeypatch.setattr(P, "main", _capture(captured))
    monkeypatch.setattr(sys, "argv", ["x", "my.xlsx", "--pause"])
    P.run()
    assert captured["argv"] == ["my.xlsx", "--pause"]
