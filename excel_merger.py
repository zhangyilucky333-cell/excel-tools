"""
Excel文件合并工具
支持合并多个Excel文件，相同名称的Sheet分别合并，自动去除重复标题行
"""
import pandas as pd
import os
from pathlib import Path
from typing import List, Dict
import openpyxl
from collections import defaultdict


class ExcelMerger:
    """Excel文件合并器"""
    
    def __init__(self, input_files: List[str], output_file: str = "merged.xlsx"):
        """
        初始化合并器
        
        Args:
            input_files: 输入的Excel文件路径列表
            output_file: 输出文件路径
        """
        self.input_files = input_files
        self.output_file = output_file
        
    def get_all_sheets_info(self) -> Dict[str, List[str]]:
        """
        获取所有文件中的Sheet信息
        
        Returns:
            字典，key为sheet名称，value为包含该sheet的文件列表
        """
        sheet_files = defaultdict(list)
        
        for file_path in self.input_files:
            try:
                excel_file = pd.ExcelFile(file_path)
                for sheet_name in excel_file.sheet_names:
                    sheet_files[sheet_name].append(file_path)
            except Exception as e:
                print(f"警告: 读取文件 '{file_path}' 失败: {str(e)}")
        
        return dict(sheet_files)
    
    def merge_sheets(self, sheet_name: str, file_list: List[str]) -> pd.DataFrame:
        """
        合并同名Sheet的数据
        
        Args:
            sheet_name: Sheet名称
            file_list: 包含该Sheet的文件列表
            
        Returns:
            合并后的DataFrame
        """
        merged_data = []
        header = None
        
        for idx, file_path in enumerate(file_list):
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                if df.empty:
                    print(f"  跳过空数据: {os.path.basename(file_path)} - {sheet_name}")
                    continue
                
                # 第一个文件，记录标题行
                if header is None:
                    header = df.columns.tolist()
                    merged_data.append(df)
                    print(f"  ✓ {os.path.basename(file_path)}: {len(df)} 行 (包含标题)")
                else:
                    # 后续文件，检查标题是否一致
                    current_header = df.columns.tolist()
                    if current_header == header:
                        # 标题一致，直接追加数据
                        merged_data.append(df)
                        print(f"  ✓ {os.path.basename(file_path)}: {len(df)} 行")
                    else:
                        # 标题不一致，尝试对齐列
                        print(f"  ⚠ {os.path.basename(file_path)}: 标题不一致，尝试对齐...")
                        # 重新排列列顺序以匹配第一个文件
                        aligned_df = pd.DataFrame(columns=header)
                        for col in header:
                            if col in df.columns:
                                aligned_df[col] = df[col].values
                        merged_data.append(aligned_df)
                        print(f"  ✓ 对齐后追加: {len(aligned_df)} 行")
                        
            except Exception as e:
                print(f"  ✗ 读取失败 {os.path.basename(file_path)} - {sheet_name}: {str(e)}")
        
        if not merged_data:
            return pd.DataFrame()
        
        # 合并所有数据
        result = pd.concat(merged_data, ignore_index=True)
        return result
    
    def copy_sheet_formatting(self, source_file: str, source_sheet_name: str,
                            target_wb, target_sheet_name: str):
        """
        从源文件复制Sheet格式
        
        Args:
            source_file: 源文件路径
            source_sheet_name: 源Sheet名称
            target_wb: 目标工作簿
            target_sheet_name: 目标Sheet名称
        """
        try:
            source_wb = openpyxl.load_workbook(source_file)
            if source_sheet_name not in source_wb.sheetnames:
                return
                
            source_ws = source_wb[source_sheet_name]
            target_ws = target_wb[target_sheet_name]
            
            # 复制列宽
            for col in source_ws.column_dimensions:
                if col in source_ws.column_dimensions:
                    target_ws.column_dimensions[col].width = source_ws.column_dimensions[col].width
            
            # 复制行高（仅第一行标题行）
            if 1 in source_ws.row_dimensions:
                target_ws.row_dimensions[1].height = source_ws.row_dimensions[1].height
                    
        except Exception as e:
            print(f"  复制格式时出错: {str(e)}")
    
    def merge_and_save(self) -> Dict[str, int]:
        """
        执行合并并保存文件
        
        Returns:
            字典，key为sheet名称，value为合并后的行数
        """
        print(f"\n开始合并 {len(self.input_files)} 个文件...")
        print("=" * 60)
        
        # 获取所有Sheet信息
        sheet_files = self.get_all_sheets_info()
        
        if not sheet_files:
            raise ValueError("未找到任何可合并的Sheet")
        
        print(f"\n找到 {len(sheet_files)} 个不同的Sheet名称")
        
        # 创建Excel写入器
        result_stats = {}
        
        with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
            for sheet_name, file_list in sorted(sheet_files.items()):
                print(f"\n正在合并 Sheet: '{sheet_name}' (来自 {len(file_list)} 个文件)")
                
                # 合并数据
                merged_df = self.merge_sheets(sheet_name, file_list)
                
                if not merged_df.empty:
                    # 写入数据
                    merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    result_stats[sheet_name] = len(merged_df)
                    
                    # 尝试复制格式（从第一个包含该sheet的文件）
                    try:
                        self.copy_sheet_formatting(
                            file_list[0], sheet_name,
                            writer.book, sheet_name
                        )
                    except:
                        pass
                    
                    print(f"  ✅ 合并完成: 共 {len(merged_df)} 行数据")
                else:
                    print(f"  ⚠ 跳过空Sheet")
        
        return result_stats
    
    def get_summary(self) -> str:
        """
        获取合并摘要信息
        
        Returns:
            摘要字符串
        """
        sheet_files = self.get_all_sheets_info()
        
        summary = f"""
合并配置摘要:
-----------------
输入文件数: {len(self.input_files)}
文件列表:
"""
        for idx, file_path in enumerate(self.input_files, 1):
            summary += f"  {idx}. {os.path.basename(file_path)}\n"
        
        summary += f"\nSheet统计:\n"
        for sheet_name, file_list in sorted(sheet_files.items()):
            summary += f"  - {sheet_name}: 出现在 {len(file_list)} 个文件中\n"
        
        summary += f"\n输出文件: {self.output_file}\n"
        
        return summary


def main():
    """命令行使用示例"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Excel文件合并工具')
    parser.add_argument('input_files', nargs='+', help='输入的Excel文件路径（可多个）')
    parser.add_argument('--output', '-o', default='merged.xlsx', help='输出文件名（默认: merged.xlsx）')
    
    args = parser.parse_args()
    
    # 创建合并器
    merger = ExcelMerger(
        input_files=args.input_files,
        output_file=args.output
    )
    
    # 显示摘要
    print(merger.get_summary())
    
    # 执行合并
    print("\n开始合并...")
    result_stats = merger.merge_and_save()
    
    print("\n" + "=" * 60)
    print("✅ 合并完成！")
    print(f"\n输出文件: {os.path.abspath(args.output)}")
    print(f"\nSheet统计:")
    for sheet_name, row_count in result_stats.items():
        print(f"  - {sheet_name}: {row_count} 行")


if __name__ == '__main__':
    main()
