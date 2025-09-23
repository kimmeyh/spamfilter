import os
import types
import yaml
import importlib

# Minimal smoke around mode routing without fully executing Outlook logic

def test_active_files_resolve_regex_by_default(monkeypatch):
    mod = importlib.import_module('withOutlookRulesYAML')
    agent = mod.OutlookSecurityAgent(test_mode=True)
    agent.set_active_mode(True)  # regex default
    assert os.path.basename(agent.active_rules_file) == 'rulesregex.yaml'
    assert os.path.basename(agent.active_safe_senders_file) == 'rules_safe_sendersregex.yaml'


def test_active_files_resolve_legacy_when_flag(monkeypatch):
    mod = importlib.import_module('withOutlookRulesYAML')
    agent = mod.OutlookSecurityAgent(test_mode=True)
    agent.set_active_mode(False)  # legacy
    assert os.path.basename(agent.active_rules_file) == 'rules.yaml'
    assert os.path.basename(agent.active_safe_senders_file) == 'rules_safe_senders.yaml'
