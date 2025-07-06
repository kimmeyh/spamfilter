import unittest
import os
import sys

# Add the directory containing withOutlookRulesCSV.py to the Python path
sys.path.append(os.path.abspath('.'))

from withOutlookRulesYAML import OutlookSecurityAgent

class TestDebug(unittest.TestCase):
    def test_debug_agent_creation(self):
        print('\nDEBUG: Testing agent creation in unittest context...')
        try:
            print('Creating OutlookSecurityAgent...')
            agent = OutlookSecurityAgent(debug_mode=True)
            print(f'SUCCESS: Agent created with email: {agent.email_address}')
            print(f'Target folders found: {len(agent.target_folders)}')
            
            print('Getting rules...')
            rules_json = agent.get_rules()
            print(f'SUCCESS: Rules retrieved, type: {type(rules_json)}')
            
        except ValueError as e:
            print(f'ValueError caught: {e}')
            if "Could not find any of the specified folders" in str(e):
                print('This is the folder detection error!')
                raise
        except Exception as e:
            print(f'Other exception: {e}')
            import traceback
            traceback.print_exc()
            raise

if __name__ == '__main__':
    unittest.main(verbosity=2)
