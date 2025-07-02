import re

def log_to_markdown(log_text):
    # 分割日志条目
    log_entries = re.split(r'#-{30,}', log_text)
    
    markdown_output = []
    
    for entry in log_entries:
        if not entry.strip():
            continue
            
        # 提取时间戳
        timestamp_match = re.search(r'\((.*?)\)', entry)
        timestamp = timestamp_match.group(1) if timestamp_match else "Unknown time"
        
        # 提取程序信息
        program_match = re.search(r'Program:\s*(.*?)\s*Code:\s*(.*?)\s*Time', entry, re.DOTALL)
        program = program_match.group(1).strip() if program_match else "Unknown program"
        code = program_match.group(2).strip() if program_match else "Unknown code"
        
        # 提取时间信息
        time_info_match = re.search(r'Time start:(.*?)\s*ending:(.*?)\s*consuming:(.*?)\n', entry, re.DOTALL)
        time_start = time_info_match.group(1).strip() if time_info_match else "Unknown start time"
        time_end = time_info_match.group(2).strip() if time_info_match else "Unknown end time"
        duration = time_info_match.group(3).strip() if time_info_match else "Unknown duration"
        
        # 提取请求和响应JSON
        request_match = re.search(r'Request JSON:\s*(.*?)\s*Response JSON:', entry, re.DOTALL)
        request_json = request_match.group(1).strip() if request_match else "No request data"
        
        response_match = re.search(r'Response JSON:\s*(.*?)\s*#', entry, re.DOTALL)
        response_json = response_match.group(1).strip() if response_match else "No response data"
        
        # 构建Markdown
        markdown_entry = f"""
## {timestamp}

- **Program**: {program}
- **Code**: {code}
- **Time Start**: {time_start}
- **Time End**: {time_end}
- **Duration**: {duration}

### Request
```
{request_json}
```

### Response
```
{response_json}
```

---
"""
        markdown_output.append(markdown_entry)
    
    return "\n".join(markdown_output)

# 示例使用
log_text = """
#--------------------------- (2025-07-02 09:38:20) ------------------------#

Program: bsft001_wf
Code: bsft001_wf

Time start:Begin Time at 2025-07-02 09:38:20

    ending:2025-07-02 09:38:20
 consuming: 0 00:00:00
Request JSON:
{"LANGUAGE":"1","USERID":"ERP","PASSWORD":"ERP","FACTORY":"XY01_ASSY","PROCSTEP":"I","CUSTOMER_CODE":" ","WO_NO":"WX-AS00-25070003","WO_TYPE":"ASSY","WO_START_DATE":"20250702","PRODUCT_ID":"ATXPTLG0001","WO_QTY_1":1,"WO_QTY_2":18214000,"N_PO_NO":" ","N_WO_NO":" ","DATE_CODE":" ","USCE_1":"WX-TK01-25270001","USCE_2":" ","USCE_3":" ","USCE_4":" ","USCE_5":" ","LOT_ID":"WX-AS00-25070003","LOT_TYPE":"E","LOT_PRIORITY":"5","MATERIAL_LOT_LIST":[],"PRD_LIST":[{"TYPE":"P","PRODUCT_ID":"SA2765A","MAINCHIP":"Y","SEQ_NUM":1,"PRODUCT_UNIT":"PCS","PART_GRP":"SA2765A","PRODUCT_QTY":1,"OPER":"A110"},{"TYPE":"P","PRODUCT_ID":"GSXPTAF0001","MAINCHIP":"N","SEQ_NUM":2,"PRODUCT_UNIT":"PCS","PART_GRP":"GSXPTAF0001","PRODUCT_QTY":1,"OPER":"A120"},{"TYPE":"P","PRODUCT_ID":"GSXPTAF0002","MAINCHIP":"N","SEQ_NUM":3,"PRODUCT_UNIT":"PCS","PART_GRP":"GSXPTAF0002","PRODUCT_QTY":1,"OPER":"A130"},{"TYPE":"M","PRODUCT_ID":"31EP0001","MAINCHIP":"N","SEQ_NUM":4,"PRODUCT_UNIT":"ML","PART_GRP":"31EP0001","PRODUCT_QTY":0.000035,"OPER":"A110"},{"TYPE":"M","PRODUCT_ID":"33CP0001","MAINCHIP":"N","SEQ_NUM":5,"PRODUCT_UNIT":"M","PART_GRP":"33CP0001","PRODUCT_QTY":0.052885,"OPER":"A310"},{"TYPE":"M","PRODUCT_ID":"34CM0001","MAINCHIP":"N","SEQ_NUM":6,"PRODUCT_UNIT":"KG","PART_GRP":"34CM0001","PRODUCT_QTY":0.000017,"OPER":"A520"},{"TYPE":"M","PRODUCT_ID":"SMXPTAF0001","MAINCHIP":"N","SEQ_NUM":7,"PRODUCT_UNIT":"EA","PART_GRP":"SMXPTAF0001","PRODUCT_QTY":1.01,"OPER":"A110"}]}

Response JSON:
{"MSG":"WIPM-P0004 : Fatal database error is occured. Please contact an administrator.","STATUSVALUE":"1"}
#------------------------------------------------------------------------------#
#--------------------------- (2025-07-02 09:50:12) ------------------------#

Program: bsft001_wf
Code: bsft001_wf

Time start:Begin Time at 2025-07-02 09:50:12

    ending:2025-07-02 09:50:12
 consuming: 0 00:00:00
Request JSON:
{"LANGUAGE":"1","USERID":"ERP","PASSWORD":"ERP","FACTORY":"XY01_ASSY","PROCSTEP":"I","CUSTOMER_CODE":" ","WO_NO":"WX-AS00-25070003","WO_TYPE":"ASSY","WO_START_DATE":"20250702","PRODUCT_ID":"ATXPTLG0001","WO_QTY_1":1,"WO_QTY_2":18214000,"N_PO_NO":" ","N_WO_NO":" ","DATE_CODE":" ","USCE_1":"WX-TK01-25270001","USCE_2":" ","USCE_3":" ","USCE_4":" ","USCE_5":" ","LOT_ID":"WX-AS00-25070003","LOT_TYPE":"E","LOT_PRIORITY":"5","MATERIAL_LOT_LIST":[],"PRD_LIST":[{"TYPE":"P","PRODUCT_ID":"SA2765A","MAINCHIP":"Y","SEQ_NUM":1,"PRODUCT_UNIT":"PCS","PART_GRP":"SA2765A","PRODUCT_QTY":1,"OPER":"A110"},{"TYPE":"P","PRODUCT_ID":"GSXPTAF0001","MAINCHIP":"N","SEQ_NUM":2,"PRODUCT_UNIT":"PCS","PART_GRP":"GSXPTAF0001","PRODUCT_QTY":1,"OPER":"A120"},{"TYPE":"P","PRODUCT_ID":"GSXPTAF0002","MAINCHIP":"N","SEQ_NUM":3,"PRODUCT_UNIT":"PCS","PART_GRP":"GSXPTAF0002","PRODUCT_QTY":1,"OPER":"A130"},{"TYPE":"M","PRODUCT_ID":"31EP0001","MAINCHIP":"N","SEQ_NUM":4,"PRODUCT_UNIT":"ML","PART_GRP":"31EP0001","PRODUCT_QTY":0.000035,"OPER":"A110"},{"TYPE":"M","PRODUCT_ID":"33CP0001","MAINCHIP":"N","SEQ_NUM":5,"PRODUCT_UNIT":"M","PART_GRP":"33CP0001","PRODUCT_QTY":0.052885,"OPER":"A310"},{"TYPE":"M","PRODUCT_ID":"34CM0001","MAINCHIP":"N","SEQ_NUM":6,"PRODUCT_UNIT":"KG","PART_GRP":"34CM0001","PRODUCT_QTY":0.000017,"OPER":"A520"},{"TYPE":"M","PRODUCT_ID":"SMXPTAF0001","MAINCHIP":"N","SEQ_NUM":7,"PRODUCT_UNIT":"EA","PART_GRP":"SMXPTAF0001","PRODUCT_QTY":1.01,"OPER":"A110"}]}

Response JSON:
{"MSG":"WIPM-P0004 : Fatal database error is occured. Please contact an administrator.","STATUSVALUE":"1"}
#------------------------------------------------------------------------------#
#--------------------------- (2025-07-02 09:52:32) ------------------------#

Program: bsft001_wf
Code: bsft001_wf

Time start:Begin Time at 2025-07-02 09:52:31

    ending:2025-07-02 09:52:32
 consuming: 0 00:00:01
Request JSON:
{"LANGUAGE":"1","USERID":"ERP","PASSWORD":"ERP","FACTORY":"XY01_ASSY","PROCSTEP":"I","CUSTOMER_CODE":" ","WO_NO":"WX-AS00-25070003","WO_TYPE":"ASSY","WO_START_DATE":"20250702","PRODUCT_ID":"ATXPTLG0001","WO_QTY_1":1,"WO_QTY_2":18214,"N_PO_NO":" ","N_WO_NO":" ","DATE_CODE":" ","USCE_1":"WX-TK01-25270001","USCE_2":" ","USCE_3":" ","USCE_4":" ","USCE_5":" ","LOT_ID":"WX-AS00-25070003","LOT_TYPE":"E","LOT_PRIORITY":"5","MATERIAL_LOT_LIST":[{"WAFER_NAME":"SA2765A","WO_LOT_ID":" ","WAFER_LOT_ID":"20250624211","WAFER_SEQ":"001","WAFER_BIN":"BIN01","WAFER_BIN_QTY":18214}],"PRD_LIST":[{"TYPE":"P","PRODUCT_ID":"SA2765A","MAINCHIP":"Y","SEQ_NUM":1,"PRODUCT_UNIT":"PCS","PART_GRP":"SA2765A","PRODUCT_QTY":1,"OPER":"A110"},{"TYPE":"P","PRODUCT_ID":"GSXPTAF0001","MAINCHIP":"N","SEQ_NUM":2,"PRODUCT_UNIT":"PCS","PART_GRP":"GSXPTAF0001","PRODUCT_QTY":1,"OPER":"A120"},{"TYPE":"P","PRODUCT_ID":"GSXPTAF0002","MAINCHIP":"N","SEQ_NUM":3,"PRODUCT_UNIT":"PCS","PART_GRP":"GSXPTAF0002","PRODUCT_QTY":1,"OPER":"A130"},{"TYPE":"M","PRODUCT_ID":"31EP0001","MAINCHIP":"N","SEQ_NUM":4,"PRODUCT_UNIT":"ML","PART_GRP":"31EP0001","PRODUCT_QTY":0.000035,"OPER":"A110"},{"TYPE":"M","PRODUCT_ID":"33CP0001","MAINCHIP":"N","SEQ_NUM":5,"PRODUCT_UNIT":"M","PART_GRP":"33CP0001","PRODUCT_QTY":0.052885,"OPER":"A310"},{"TYPE":"M","PRODUCT_ID":"34CM0001","MAINCHIP":"N","SEQ_NUM":6,"PRODUCT_UNIT":"KG","PART_GRP":"34CM0001","PRODUCT_QTY":0.000017,"OPER":"A520"},{"TYPE":"M","PRODUCT_ID":"SMXPTAF0001","MAINCHIP":"N","SEQ_NUM":7,"PRODUCT_UNIT":"EA","PART_GRP":"SMXPTAF0001","PRODUCT_QTY":1.01,"OPER":"A110"}]}

Response JSON:
{"MSG":"This service is successful","STATUSVALUE":"0"}
#------------------------------------------------------------------------------#
#--------------------------- (2025-07-02 09:54:47) ------------------------#

Program: axmm200
Code: apmm100

Time start:Begin Time at 2025-07-02 09:54:47

    ending:2025-07-02 09:54:47
 consuming: 0 00:00:00
Request JSON:
{"LANGUAGE":"1","USERID":"ERP","PASSWORD":"ERP","FACTORY":"XY01_ASSY","PROCSTEP":"I","CUST_ID":"000049","CUST_DESC":"测试客户接口"}

Response JSON:
{"MSG":"This service is successful","STATUSVALUE":"0"}
#------------------------------------------------------------------------------#
#--------------------------- (2025-07-02 09:54:47) ------------------------#

Program: axmm200
Code: apmm100

Time start:Begin Time at 2025-07-02 09:54:47

    ending:2025-07-02 09:54:47
 consuming: 0 00:00:00
Request JSON:
{"LANGUAGE":"1","USERID":"ERP","PASSWORD":"ERP","FACTORY":"XY01_GS","PROCSTEP":"I","CUST_ID":"000049","CUST_DESC":"测试客户接口"}

Response JSON:
{"MSG":"This service is successful","STATUSVALUE":"0"}
#------------------------------------------------------------------------------#
#--------------------------- (2025-07-02 09:54:47) ------------------------#

Program: axmm200
Code: apmm100

Time start:Begin Time at 2025-07-02 09:54:47

    ending:2025-07-02 09:54:47
 consuming: 0 00:00:00
Request JSON:
{"LANGUAGE":"1","USERID":"ERP","PASSWORD":"ERP","FACTORY":"XY01_SMT","PROCSTEP":"I","CUST_ID":"000049","CUST_DESC":"测试客户接口"}

Response JSON:
{"MSG":"This service is successful","STATUSVALUE":"0"}
#------------------------------------------------------------------------------#
#--------------------------- (2025-07-02 09:54:47) ------------------------#

Program: axmm200
Code: apmm100

Time start:Begin Time at 2025-07-02 09:54:47

    ending:2025-07-02 09:54:47
 consuming: 0 00:00:00
Request JSON:
{"LANGUAGE":"1","USERID":"ERP","PASSWORD":"ERP","FACTORY":"XY01_FT","PROCSTEP":"I","CUST_ID":"000049","CUST_DESC":"测试客户接口"}

Response JSON:
{"MSG":"This service is successful","STATUSVALUE":"0"}
#------------------------------------------------------------------------------#
#--------------------------- (2025-07-02 09:54:47) ------------------------#

Program: axmm200
Code: apmm100

Time start:Begin Time at 2025-07-02 09:54:47

    ending:2025-07-02 09:54:47
 consuming: 0 00:00:00
Request JSON:
{"LANGUAGE":"1","USERID":"ERP","PASSWORD":"ERP","FACTORY":"XY01_CP","PROCSTEP":"I","CUST_ID":"000049","CUST_DESC":"测试客户接口"}

Response JSON:
{"MSG":"This service is successful","STATUSVALUE":"0"}
#------------------------------------------------------------------------------#

"""
markdown_result = log_to_markdown(log_text)
print(markdown_result)
