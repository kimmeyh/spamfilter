import sys
import os
sys.path.append(os.path.abspath("."))

def test_debug_outlook():
    print("\nDEBUG: Running in pytest context")
    print(f"Working directory: {os.getcwd()}")
    print(f"Python path: {sys.path[:3]}...")
    
    try:
        from withOutlookRulesYAML import OutlookSecurityAgent
        print("Import successful")
        
        agent = OutlookSecurityAgent(debug_mode=False)
        print(f"Agent created with {len(agent.target_folders)} folders")
        assert len(agent.target_folders) > 0
        
    except Exception as e:
        print(f"Error in test: {e}")
        import traceback
        traceback.print_exc()
        raise
