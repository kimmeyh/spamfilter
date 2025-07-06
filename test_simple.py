"""
Simplified test to isolate the issue
"""

import pytest
from withOutlookRulesYAML import OutlookSecurityAgent


def test_simple_agent_creation():
    """Simple test of agent creation"""
    print("\n=== Simple Agent Creation Test ===")
    
    # Create agent in test mode - this should pass even without Outlook folders
    agent = OutlookSecurityAgent(debug_mode=True, test_mode=True)
    print(f"SUCCESS: Created agent in test mode")
    
    # In test mode, we don't require folders to exist
    print(f"Agent created with {len(agent.target_folders)} folders (test mode allows 0)")
    assert agent is not None


class TestSimpleClass:
    """Test class with fixture"""
    
    @pytest.fixture(autouse=True)
    def setup(self):
        """Setup fixture"""
        print("\n=== Class Setup ===")
        # Create agent in test mode - this should pass even without Outlook folders
        self.agent = OutlookSecurityAgent(debug_mode=True, test_mode=True)
        print(f"SUCCESS: Created agent in fixture (test mode)")
    
    def test_agent_in_class(self):
        """Test agent creation in class"""
        print("\n=== Class Test Method ===")
        assert self.agent is not None
        # In test mode, we don't require folders to exist
        print(f"SUCCESS: Agent created in test mode with {len(self.agent.target_folders)} folders")


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])
