#!/usr/bin/env python3
"""
Test script for second-pass email processing implementation
Tests the new _get_emails_from_folder method and second-pass processing logic
"""

import ast
import inspect

def test_helper_method_exists():
    """Test that the _get_emails_from_folder helper method exists"""
    try:
        with open('withOutlookRulesYAML.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Check if the method exists
        if '_get_emails_from_folder' in content:
            print("✓ _get_emails_from_folder method found")
            return True
        else:
            print("✗ _get_emails_from_folder method not found")
            return False
    except Exception as e:
        print(f"✗ Error reading file: {e}")
        return False

def test_second_pass_logic():
    """Test that second-pass processing logic has been added"""
    try:
        with open('withOutlookRulesYAML.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Check for second-pass processing keywords
        second_pass_indicators = [
            'Second-pass processing:',
            'second_pass_emails',
            'second_pass_added_info',
            'Starting second-pass email processing',
            'Second-pass: Found',
            'Second-pass Processing Summary'
        ]
        
        found_indicators = []
        for indicator in second_pass_indicators:
            if indicator in content:
                found_indicators.append(indicator)
        
        print(f"✓ Found {len(found_indicators)}/{len(second_pass_indicators)} second-pass indicators")
        
        if len(found_indicators) >= 4:  # At least 4 out of 6 indicators should be present
            print("✓ Second-pass processing logic appears to be implemented")
            return True
        else:
            print("✗ Second-pass processing logic incomplete")
            return False
            
    except Exception as e:
        print(f"✗ Error reading file: {e}")
        return False

def test_method_signature():
    """Test that the _get_emails_from_folder method has correct signature"""
    try:
        with open('withOutlookRulesYAML.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Parse the AST to find the method
        tree = ast.parse(content)
        
        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef) and node.name == '_get_emails_from_folder':
                # Check method signature
                args = [arg.arg for arg in node.args.args]
                expected_args = ['self', 'folder', 'days_back']
                
                if args == expected_args:
                    print(f"✓ Method signature correct: {args}")
                    return True
                else:
                    print(f"✗ Method signature incorrect. Expected: {expected_args}, Got: {args}")
                    return False
        
        print("✗ Method _get_emails_from_folder not found in AST")
        return False
        
    except Exception as e:
        print(f"✗ Error parsing file: {e}")
        return False

def test_after_prompt_update_rules():
    """Test that second-pass processing is placed after prompt_update_rules"""
    try:
        with open('withOutlookRulesYAML.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Find positions of key text
        prompt_update_pos = content.find('self.prompt_update_rules(')
        second_pass_pos = content.find('Second-pass processing:')
        
        if prompt_update_pos == -1:
            print("✗ prompt_update_rules call not found")
            return False
        
        if second_pass_pos == -1:
            print("✗ Second-pass processing not found")
            return False
        
        if second_pass_pos > prompt_update_pos:
            print("✓ Second-pass processing correctly placed after prompt_update_rules")
            return True
        else:
            print("✗ Second-pass processing not in correct position")
            return False
            
    except Exception as e:
        print(f"✗ Error reading file: {e}")
        return False

def main():
    """Run all tests for second-pass implementation"""
    print("Testing second-pass email processing implementation")
    print("=" * 60)
    
    tests = [
        test_helper_method_exists,
        test_method_signature,
        test_second_pass_logic,
        test_after_prompt_update_rules
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        print(f"\nRunning {test.__name__}...")
        if test():
            passed += 1
        else:
            print(f"❌ {test.__name__} FAILED")
    
    print("\n" + "=" * 60)
    print(f"Test Results: {passed} passed, {total - passed} failed")
    
    if passed == total:
        print("✅ All second-pass implementation tests PASSED!")
        return True
    else:
        print("❌ Some tests failed. Please check the implementation.")
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
