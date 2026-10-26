from __future__ import annotations


class SheetPilotError(Exception):
    code = "SHEETPILOT_ERROR"

    def __init__(self, message: str, *, detail: dict | None = None):
        super().__init__(message)
        self.message = message
        self.detail = detail or {}


class FileRejectedError(SheetPilotError):
    code = "FILE_REJECTED"


class WorkbookStructureError(SheetPilotError):
    code = "WORKBOOK_STRUCTURE_ERROR"


class IntentValidationError(SheetPilotError):
    code = "INTENT_VALIDATION_ERROR"


class DuplicateKeyError(SheetPilotError):
    code = "DUPLICATE_SOURCE_KEY"


class ExecutionError(SheetPilotError):
    code = "EXECUTION_ERROR"


class VerificationError(SheetPilotError):
    code = "VERIFICATION_FAILED"


class LLMError(SheetPilotError):
    code = "LLM_ERROR"

