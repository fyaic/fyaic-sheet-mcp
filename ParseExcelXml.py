"""
Excel 表格快速读取工具
使用方法: python ParseExcelXml.py "Excel文件路径" [工作表索引] [最大行数]
"""

import xml.etree.ElementTree as ET
import json
import sys
import os
from pathlib import Path

# 设置 UTF-8 编码输出
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())


def parse_excel_xml(xlsx_path, sheet_index=1, max_rows=300):
    """
    解析 Excel 文件的 XML 内容

    参数:
        xlsx_path: Excel 文件路径
        sheet_index: 工作表索引(从1开始,默认1)
        max_rows: 最大读取行数(默认300)
    """

    xlsx_path = Path(xlsx_path)

    # 检查文件是否存在
    if not xlsx_path.exists():
        return {"error": f"文件不存在: {xlsx_path}"}

    # 尝试从 _xml 文件夹读取
    xml_dir = xlsx_path.parent / f"{xlsx_path.stem}_xml"

    if not xml_dir.exists():
        return {"error": f"未找到 XML 文件夹: {xml_dir}\n请先解压 .xlsx 文件或使用 ReadSheetAsXml.ps1 导出"}

    # 读取共享字符串
    ss_path = xml_dir / "sharedStrings.xml"
    if not ss_path.exists():
        return {"error": f"未找到 sharedStrings.xml"}

    try:
        ss_tree = ET.parse(ss_path)
        ss_root = ss_tree.getroot()
        shared_strings = [
            si.text for si in ss_root.findall(
                './/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si/'
                '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'
            )
        ]
    except Exception as e:
        return {"error": f"解析 sharedStrings.xml 失败: {str(e)}"}

    # 读取指定工作表
    sheet_path = xml_dir / f"sheet{sheet_index}.xml"
    if not sheet_path.exists():
        return {"error": f"未找到 sheet{sheet_index}.xml"}

    try:
        sheet_tree = ET.parse(sheet_path)
        sheet_root = sheet_tree.getroot()

        # 命名空间
        ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

        # 解析单元格数据
        rows_data = {}

        for row in sheet_root.findall('.//row', ns):
            row_num = int(row.get('r'))
            if row_num > max_rows:
                break

            cells = {}
            for c in row.findall('c', ns):
                addr = c.get('r')
                cell_type = c.get('t', '')

                v_elem = c.find('v', ns)
                if v_elem is not None:
                    raw_value = v_elem.text
                    if raw_value:
                        if cell_type == 's':  # 共享字符串
                            idx = int(raw_value)
                            if 0 <= idx < len(shared_strings):
                                value = shared_strings[idx]
                            else:
                                value = raw_value
                        else:
                            value = raw_value

                        cells[addr] = value

            if cells:
                rows_data[row_num] = cells

        return {
            "success": True,
            "file": str(xlsx_path),
            "sheet_index": sheet_index,
            "total_rows": len(rows_data),
            "data": rows_data
        }

    except Exception as e:
        return {"error": f"解析 sheet{sheet_index}.xml 失败: {str(e)}"}


def main():
    """命令行入口"""
    if len(sys.argv) < 2:
        print("使用方法: python read_excel.py <Excel文件路径> [工作表索引] [最大行数]")
        print("示例: python read_excel.py 'data.xlsx' 1 500")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    sheet_index = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    max_rows = int(sys.argv[3]) if len(sys.argv) > 3 else 300

    result = parse_excel_xml(xlsx_path, sheet_index, max_rows)

    # 输出 JSON
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
