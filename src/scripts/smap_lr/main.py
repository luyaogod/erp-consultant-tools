from .core.handers import CopyCloumnHandler
from .core.processor import Processor
from pydantic import BaseModel, Field

class ToolConfig(BaseModel):
    Sheet1_name: str = Field(..., description="待处理Sheet名")
    Sheet2_name: str = Field(..., description="参照表Sheet名")
    Sheet3_name: str = Field(..., description="处理结果Sheet名")
    direction: bool = Field(..., description="匹配方向，False表示反转匹配方向")
    row1: int = Field(..., description="待处理Sheet匹配行", gt=0)
    row2: int = Field(..., description="参照表Sheet匹配行", gt=0)
    fp: str = Field(..., description="待处理Excel文件路径")
    max_num: int = Field(2000, description="CopyCloumnHandler处理器最大copy行数", gt=0)

    class Config:
        # 添加额外的配置，例如字段别名等
        extra = "forbid"  # 禁止额外字段
        anystr_strip_whitespace = True  # 自动去除字符串两端空格

def run_tool_copy_clomun(
    config: ToolConfig
):
    p = Processor(
        file_path=config.fp,
        sheet1_name=config.Sheet1_name,
        sheet2_name=config.Sheet2_name,
        sheet3_name=config.Sheet3_name,
    )
    h = CopyCloumnHandler(config.max_num)
    p.process(
        direction=config.direction,
        row1=config.row1,
        row2=config.row2,
        on_match=h.my_on_match,
        on_nomatch=h.my_on_nomatch,
    )