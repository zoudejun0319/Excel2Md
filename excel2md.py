"""
Excel转Markdown工具
支持将Excel文件转换为Markdown表格格式
"""

import pandas as pd
import argparse
import os
import sys
from pathlib import Path


class Excel2Markdown:
    """Excel转Markdown转换器"""

    def __init__(self, excel_path, output_path=None, sheet_name=None):
        """
        初始化转换器

        Args:
            excel_path: Excel文件路径
            output_path: 输出Markdown文件路径（可选）
            sheet_name: 工作表名称（可选，默认转换所有工作表）
        """
        self.excel_path = Path(excel_path)
        self.output_path = Path(output_path) if output_path else None
        self.sheet_name = sheet_name

    def read_excel(self):
        """
        读取Excel文件

        Returns:
            DataFrame字典或单个DataFrame
        """
        if not self.excel_path.exists():
            raise FileNotFoundError(f"文件不存在: {self.excel_path}")

        try:
            if self.sheet_name:
                # 读取指定工作表
                df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name)
                return {self.sheet_name: df}
            else:
                # 读取所有工作表
                dfs = pd.read_excel(self.excel_path, sheet_name=None)
                return dfs
        except Exception as e:
            raise Exception(f"读取Excel文件失败: {str(e)}")

    @staticmethod
    def dataframe_to_markdown(df, table_name=""):
        """
        将DataFrame转换为Markdown表格

        Args:
            df: pandas DataFrame
            table_name: 表格名称/标题

        Returns:
            Markdown格式的字符串
        """
        if df.empty:
            return f"# {table_name}\n\n表格为空\n"

        # 处理NaN值
        df = df.fillna("")

        # 生成表头
        markdown_lines = []

        if table_name:
            markdown_lines.append(f"## {table_name}")
            markdown_lines.append("")

        # 表头
        headers = df.columns.tolist()
        markdown_lines.append("| " + " | ".join(str(h) for h in headers) + " |")

        # 分隔线
        markdown_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")

        # 数据行
        for _, row in df.iterrows():
            row_data = " | ".join(str(val) for val in row.tolist())
            markdown_lines.append(f"| {row_data} |")

        return "\n".join(markdown_lines)

    def convert(self):
        """
        执行转换

        Returns:
            Markdown字符串
        """
        dfs = self.read_excel()

        if isinstance(dfs, dict):
            # 多个工作表
            markdown_sections = []
            for sheet_name, df in dfs.items():
                section = self.dataframe_to_markdown(df, sheet_name)
                markdown_sections.append(section)
                markdown_sections.append("\n\n---\n\n")
            markdown = "\n".join(markdown_sections)
        else:
            # 单个工作表
            markdown = self.dataframe_to_markdown(dfs)

        return markdown

    def save(self, markdown_content):
        """
        保存Markdown文件

        Args:
            markdown_content: Markdown内容
        """
        if self.output_path is None:
            # 默认输出路径
            self.output_path = self.excel_path.with_suffix('.md')

        with open(self.output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)

        print(f"✓ 已保存到: {self.output_path}")

    def run(self):
        """执行完整的转换流程"""
        try:
            print(f"正在读取: {self.excel_path}")
            markdown = self.convert()
            self.save(markdown)
            print("✓ 转换完成!")
            return True
        except Exception as e:
            print(f"✗ 转换失败: {str(e)}", file=sys.stderr)
            return False


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(
        description='将Excel文件转换为Markdown格式',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s data.xlsx                    # 转换为data.md
  %(prog)s data.xlsx -o output.md       # 指定输出文件
  %(prog)s data.xlsx -s Sheet1          # 只转换Sheet1工作表
        """
    )

    parser.add_argument(
        'excel_file',
        help='Excel文件路径 (.xlsx, .xls)'
    )

    parser.add_argument(
        '-o', '--output',
        help='输出Markdown文件路径（默认与Excel文件同名）'
    )

    parser.add_argument(
        '-s', '--sheet',
        help='指定要转换的工作表名称（默认转换所有工作表）'
    )

    args = parser.parse_args()

    # 执行转换
    converter = Excel2Markdown(
        excel_path=args.excel_file,
        output_path=args.output,
        sheet_name=args.sheet
    )

    return converter.run()


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
