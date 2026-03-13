import importlib
import sys
import types
import unittest
from pathlib import Path


def _bootstrap_outlook_package() -> None:
    if "outlook" in sys.modules:
        return

    root = Path(__file__).resolve().parents[1]
    pkg = types.ModuleType("outlook")
    pkg.__path__ = [str(root)]
    sys.modules["outlook"] = pkg


_bootstrap_outlook_package()
validation = importlib.import_module("outlook.validation")
exceptions = importlib.import_module("outlook.exceptions")


class CustomExceptionsTests(unittest.TestCase):
    def test_custom_exceptions_keep_builtin_compatibility(self) -> None:
        self.assertTrue(issubclass(exceptions.EmailValidationError, ValueError))
        self.assertTrue(issubclass(exceptions.PathValidationError, ValueError))
        self.assertTrue(
            issubclass(exceptions.OutlookConnectionError, RuntimeError)
        )

    def test_invalid_email_raises_email_validation_error(self) -> None:
        with self.assertRaises(exceptions.EmailValidationError):
            validation.validate_email("invalid-email")

    def test_missing_path_raises_path_validation_error(self) -> None:
        with self.assertRaises(exceptions.PathValidationError):
            validation.validate_paths("/definitely/missing/path/for/outlook-nav-test")
