import pandas as pd
from rapidfuzz import process, fuzz
from tqdm import tqdm
import argparse
import sys
from multiprocessing import Pool, cpu_count
import numpy as np


def find_best_match(query, choices, threshold=75):
    """
    在choices中找到与query最匹配的项
    返回匹配索引和匹配分数
    """
    # 使用rapidfuzz的process.extractOne进行快速匹配
    result = process.extractOne(
        query,
        choices,
        scorer=fuzz.token_set_ratio,
        score_cutoff=threshold
    )
    if result:
        match_str, score, idx = result
        return idx, score
    return None, 0


def process_chunk(args):
    """
    处理数据块的函数，用于多进程
    """
    alm_chunk, issue_subjects, threshold, issues_df = args
    results = []

    for _, alm_row in alm_chunk.iterrows():
        nonconformity = str(alm_row['不符合现象'])
        match_idx, match_score = find_best_match(
            nonconformity, issue_subjects, threshold)

        row_data = alm_row.to_dict()
        row_data['匹配分数'] = match_score

        if match_idx is not None:
            matched_issue = issues_df.iloc[match_idx]
            for col in issues_df.columns:
                row_data[f"Issues_{col}"] = matched_issue[col]
            row_data['主题匹配结果'] = matched_issue['主题']
            row_data['匹配状态'] = '成功匹配'
        else:
            for col in issues_df.columns:
                row_data[f"Issues_{col}"] = ""
            row_data['主题匹配结果'] = ""
            row_data['匹配状态'] = '未找到匹配'

        results.append(row_data)

    return results


def merge_alm_issues(alm_path, issues_path, output_path, alm_encoding='ANSI', issues_encoding='ANSI', threshold=75, n_workers=None):
    """
    合并ALM和Issues表格基于模糊匹配（优化版）
    参数:
    alm_path -- ALM文件路径
    issues_path -- Issues文件路径
    output_path -- 输出文件路径
    alm_encoding -- ALM文件编码 (默认: 'ANSI')
    issues_encoding -- Issues文件编码 (默认: 'ANSI')
    threshold -- 模糊匹配阈值 (默认: 75)
    n_workers -- 并行工作进程数 (默认: CPU核心数)
    """
    # 读取两个CSV文件
    alm_df = pd.read_csv(alm_path, encoding=alm_encoding)
    issues_df = pd.read_csv(issues_path, encoding=issues_encoding)

    # 为issues表创建主题列表用于匹配
    issue_subjects = issues_df['主题'].tolist()

    # 设置工作进程数
    if n_workers is None:
        n_workers = min(cpu_count(), 8)  # 最多使用8个进程

    # 将ALM数据分成多个块
    chunk_size = max(100, len(alm_df) // (n_workers * 4))  # 每个块至少100行
    chunks = [alm_df[i:i+chunk_size]
              for i in range(0, len(alm_df), chunk_size)]

    # 准备多进程参数
    pool_args = [(chunk, issue_subjects, threshold, issues_df)
                 for chunk in chunks]

    # 使用多进程处理
    merged_data = []
    with Pool(processes=n_workers) as pool:
        # 使用tqdm显示进度
        with tqdm(total=len(chunks), desc="处理进度") as pbar:
            for result in pool.imap(process_chunk, pool_args):
                merged_data.extend(result)
                pbar.update(1)

    # 创建合并后的DataFrame
    merged_df = pd.DataFrame(merged_data)

    # 调整列顺序以便阅读
    core_columns = ['编号', '匹配状态', '匹配分数', '不符合现象', '主题匹配结果']
    other_columns = [
        col for col in merged_df.columns if col not in core_columns]
    merged_df = merged_df[core_columns + other_columns]

    # 保存合并结果
    merged_df.to_csv(output_path, index=False, encoding='utf_8_sig')

    # 打印统计信息
    total_rows = len(merged_df)
    matched_rows = merged_df['匹配状态'].eq('成功匹配').sum()
    print(f"合并完成! 结果已保存至: {output_path}")
    print(f"总行数: {total_rows}")
    print(f"成功匹配行数: {matched_rows} ({matched_rows/total_rows:.1%})")
    print(f"使用进程数: {n_workers}")

    return merged_df


def main():
    # 创建命令行参数解析器
    parser = argparse.ArgumentParser(description='ALM和Issues表格合并工具(优化版)')

    # 添加命令行参数
    parser.add_argument('-a', '--alm', required=True, help='ALM文件路径')
    parser.add_argument('-i', '--issues', required=True, help='Issues文件路径')
    parser.add_argument('-o', '--output', required=True, help='输出文件路径')
    parser.add_argument('--alm-encoding', default='ANSI',
                        help='ALM文件编码 (默认: ANSI)')
    parser.add_argument('--issues-encoding', default='ANSI',
                        help='Issues文件编码 (默认: ANSI)')
    parser.add_argument('-t', '--threshold', type=int,
                        default=75, help='模糊匹配阈值 (默认: 75)')
    parser.add_argument('-n', '--n-workers', type=int, default=None,
                        help='并行工作进程数 (默认: CPU核心数，最多8个)')

    # 解析命令行参数
    args = parser.parse_args()

    # 打印参数信息
    print("ALM和Issues表格合并工具(优化版)")
    print("=" * 50)
    print(f"ALM文件: {args.alm}")
    print(f"Issues文件: {args.issues}")
    print(f"输出文件: {args.output}")
    print(f"ALM编码: {args.alm_encoding}")
    print(f"Issues编码: {args.issues_encoding}")
    print(f"匹配阈值: {args.threshold}")
    print(f"工作进程数: {args.n_workers if args.n_workers else '自动(CPU核心数)'}")
    print("=" * 50)

    # 执行合并操作
    try:
        result = merge_alm_issues(
            alm_path=args.alm,
            issues_path=args.issues,
            output_path=args.output,
            alm_encoding=args.alm_encoding,
            issues_encoding=args.issues_encoding,
            threshold=args.threshold,
            n_workers=args.n_workers
        )
        print("\n[SUCCESS] 合并操作成功完成！")
    except Exception as e:
        print(f"\n[ERROR] 合并操作失败: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
