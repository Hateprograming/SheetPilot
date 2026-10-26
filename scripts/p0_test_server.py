"""Isolated browser-test server. Uses deterministic model substitutes, no network model calls."""
import os
from dataclasses import replace
from pathlib import Path

os.environ['SHEETPILOT_WEB_DATA_DIR'] = str(Path(__file__).resolve().parents[1]/'.pytest-p0-browser-data')

from sheetpilot import web_api
from sheetpilot.llm import FakeLLMClient
from sheetpilot.schemas import TaskEnvelope, LookupTaskSpec
from sheetpilot.service import SheetPilotService

spec = TaskEnvelope(goal='按客户编号匹配客户等级', spec=LookupTaskSpec(
    main_sheet='订单',source_sheet='客户',main_key_column='客户编号',source_key_column='客户编号',
    return_columns=['客户等级'], confidence=1))
web_api.settings = replace(web_api.settings, api_key='', base_url='', model='test-double')
web_api.service = SheetPilotService(web_api.settings, FakeLLMClient(spec))

if __name__ == '__main__':
    import uvicorn
    uvicorn.run(web_api.app, host='127.0.0.1', port=8769)
