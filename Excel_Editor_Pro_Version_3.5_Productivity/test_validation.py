
import sys
import unittest
from PyQt5.QtWidgets import QApplication

try:
    from DataValidation_ import ValidationRule, ValidationManager
    print("[OK] DataValidation_ module imported successfully")
except ImportError as e:
    print(f"[FAIL] Failed to import DataValidation_ module: {e}")
    sys.exit(1)

pass

class TestValidation(unittest.TestCase):
    def test_numeric_rule(self):
        rule = ValidationRule(ValidationRule.TYPE_NUMBER, operator='>', value=10)
        
        valid, msg = rule.validate("15")
        self.assertTrue(valid)
        
        valid, msg = rule.validate("5")
        self.assertFalse(valid)
        self.assertIn("must be > 10", msg)
        
        valid, msg = rule.validate("abc")
        self.assertFalse(valid)
        self.assertEqual(msg, "Invalid data type")

    def test_text_length_rule(self):
        rule = ValidationRule(ValidationRule.TYPE_TEXT, operator='<', value=5)
        
        valid, msg = rule.validate("Hi")
        self.assertTrue(valid)
        
        valid, msg = rule.validate("Hello World")
        self.assertFalse(valid)
        self.assertIn("length must be < 5", msg)

    def test_list_rule(self):
        rule = ValidationRule(ValidationRule.TYPE_LIST, options=['A', 'B', 'C'])
        
        valid, msg = rule.validate("A")
        self.assertTrue(valid)
        
        valid, msg = rule.validate("D")
        self.assertFalse(valid)
        self.assertIn("one of: A, B, C", msg)

    def test_manager(self):
        manager = ValidationManager()
        rule = ValidationRule(ValidationRule.TYPE_NUMBER, operator='>', value=0)
        manager.add_rule(0, rule)
        
        valid, msg = manager.validate_cell(0, "100")
        self.assertTrue(valid)
        
        # No rule for col 1
        valid, msg = manager.validate_cell(1, "-50")
        self.assertTrue(valid)

if __name__ == '__main__':
    # Create QApplication for potential Qt dependence (though rules are logic only)
    app = QApplication(sys.argv) 
    unittest.main()
