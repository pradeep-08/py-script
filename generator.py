import os
import openpyxl

MAPPING_RULES = [
    (["voltage to 12v", "battery voltage", "power up"], ["Set_Bat_Vol(12);", "Read_MEC();"]),
    (["mec equal to zero", "read mec"], ["Read_MEC();"]),
    (["extended diagnostic session", "extended session"], ["Extended_Session_Phy();"]),
    (["tester present"], ["Tester_Present_Per_Start_Fun(2000);"]),
    (["disable communication"], ["Dis_Comm_phy();"]),
    (["security level 03", "security level 3"], ["Security_Unlock_With_Key_lvl_3_Phy();"]),
    (["security level 01", "security level 1"], ["Security_Unlock_With_Key_lvl_1_Phy();"]),
    (["security level 05", "security level 5"], ["Security_Unlock_With_Key_lvl_5_Phy();"]),
    (["security level 09", "security level 9"], ["Security_Unlock_With_Key_lvl_9_Phy();"]),
    (["security level 11"], ["Security_Unlock_With_Key_lvl_11_Phy();"]),
    (["security level 13"], ["Security_Unlock_With_Key_lvl_13_Phy();"]),
    (["security level 15"], ["Security_Unlock_With_Key_lvl_15_Phy();"]),
    (["service id 31 and sub function 01", "rid 21e"], ["Write_RID_21E();"]),
    (["ecu reset", "hard reset"], ["Hard_Reset_Phy();"]),
    (["default session"], ["Default_Session_Phy();"]),
    (["normal mode", "back to normal"], ["Back_To_Normal();"]),
    (["power mode to off"], ["@IgnitionSwitch=0;", "Read_PowerMode();"]),
    (["power mode to acc"], ["@IgnitionSwitch=1;", "Read_PowerMode();"]),
    (["power mode to run"], ["@IgnitionSwitch=2;", "Read_PowerMode();"]),
    (["power mode to start"], ["@IgnitionSwitch=3;", "Read_PowerMode();"])
]

def map_step_to_capl(description):
    desc_lower = description.lower()
    mapped_lines = []
    for keywords, capl_codes in MAPPING_RULES:
        if any(kw in desc_lower for kw in keywords):
            for code in capl_codes:
                if code not in mapped_lines:
                    mapped_lines.append(code)
    return mapped_lines

def extract_steps_from_workbook(excel_path):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = wb.active
    
    desc_col = 2
    step_col = 1
    
    # Try to dynamically find columns
    for row in sheet.iter_rows(min_row=1, max_row=5):
        for cell in row:
            if cell.value:
                val = str(cell.value).lower().strip()
                if "description" in val:
                    desc_col = cell.column
                elif val == "step":
                    step_col = cell.column

    steps = []
    unmatched_rows = []
    
    total = 0
    mapped = 0

    for row_idx in range(2, sheet.max_row + 1):
        desc = sheet.cell(row=row_idx, column=desc_col).value
        step_name = sheet.cell(row=row_idx, column=step_col).value
        
        if not step_name:
            step_name = f"Step {row_idx - 1}"
            
        if desc:
            total += 1
            capl_commands = map_step_to_capl(str(desc))
            
            if capl_commands:
                mapped += 1
                steps.append((step_name, str(desc), capl_commands))
            else:
                unmatched_rows.append(f"{step_name}: {str(desc)[:50]}...")
                steps.append((step_name, str(desc), [f"// TODO: Unmapped logic for '{str(desc)[:30]}'"]))

    # If it's totally empty or we couldn't parse, let's at least not break
    unmatched = total - mapped
    
    return steps, total, mapped, unmatched, unmatched_rows

def build_testcase_from_steps(testcase_name, steps):
    code = f"testcase {testcase_name}()\n{{\n"
    code += '  char logPath[256] = "C:\\\\Users\\\\Logs\\\\test.asc";\n'
    code += '  setLogFileName("Test Logs",logPath);\n'
    code += '  startLogging("Test Logs");\n\n'
    
    for i, (step_name, desc, capl_commands) in enumerate(steps, 1):
        code += f"  /*-------------------------------Test Step {i} - {step_name}--------------------------------------*/\n"
        # We replace newlines with space to keep the comment clean
        clean_desc = desc.replace('\n', ' | ')
        code += f"  /* ACTION: {clean_desc} */\n"
        for capl in capl_commands:
            code += f"  {capl}\n"
        code += "\n"
        
    code += "  stopLogging(\"Test Logs\");\n"
    code += "}\n"
    return code

def generate_can_from_excel_with_master(excel_path, master_can_path, output_dir):
    try:
        # 1. Parse Excel
        steps, total, mapped, unmatched, unmatched_arr = extract_steps_from_workbook(excel_path)
        
        # 2. Build the exact testcase code
        base_name = os.path.basename(excel_path).replace('.xlsx', '').replace('.xls', '')
        # testcase names must be valid C identifiers
        safe_tc_name = "TC_Generated_" + "".join(c if c.isalnum() else "_" for c in base_name)
        testcase_code = build_testcase_from_steps(safe_tc_name, steps)
        
        # 3. Combine with master (we extract the exact includes/variables from your master)
        final_code = "/* Generated CAPL code string here */\n"
        final_code += f"/* Source: {os.path.basename(excel_path)} */\n"
        
        if os.path.exists(master_can_path):
            # Just read the first lines (includes, variables) from the master snippet
            with open(master_can_path, "r", encoding="utf-8") as m_file:
                content = m_file.read()
                # find the first testcase and extract everything before it
                idx = content.find("testcase ")
                if idx != -1:
                    final_code += content[:idx]
        
        # Append our completely customized testcase
        final_code += "\n\n" + testcase_code
        
        # Output saving logic if we wanted to save it locally
        output_path = os.path.join(output_dir, f"{safe_tc_name}.can")
        with open(output_path, "w", encoding="utf-8") as out:
            out.write(final_code)

        return {
            "total": total,
            "mapped": mapped,
            "unmatched": unmatched,
            "unmatchedSteps": unmatched_arr,
            "previewCode": final_code
        }
    except Exception as e:
        print(f"Error during generation: {e}")
        return {
            "total": 0,
            "mapped": 0,
            "unmatched": 0,
            "unmatchedSteps": [str(e)],
            "previewCode": f"// Generation failed: {str(e)}"
        }
