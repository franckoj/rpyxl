"""Compatibility import path for data validation models."""
from .._models import DataValidation, DataValidations

DataValidationList = DataValidations

__all__ = ['DataValidation', 'DataValidations', 'DataValidationList']
