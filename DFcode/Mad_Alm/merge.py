import pandas as pd
from fuzzywuzzy import fuzz
from tqdm import tqdm
import argparse
import sys


def merge_alm_issues(alm_path, issues_path, output_path, alm_encoding='ANSI', issues_encoding='ANSI', threshold=75):
    """
    合并ALM和Issues表格基于模糊匹配
    参数:
    alm_path -- ALM文件路径
    issues_path -- Issues文件路径
    output_path -- 输出文件路径
    alm_encoding -- ALM文件编码 (默认: 'ANSI')
    issues_encoding -- Issues文件编码 (默认: 'ANSI')
    threshold -- 模糊匹配阈值 (默认: 75)
    """
    # 读取两个CSV文件
    alm_df = pd.read_csv(alm_path, encoding=alm_encoding)
    issues_df = pd.read_csv(issues_path, encoding=issues_encoding)

    def find_best_match(query, choices, threshold=75):
        """
        在choices中找到与query最匹配的项
        返回匹配索引和匹配分数
        """
        best_match, best_score = None, 0
        for i, choice in enumerate(choices):
            # 使用token_set_ratio算法，对词序和重复词不敏感
            score = fuzz.token_set_ratio(query, choice)
            if score > best_score:
                best_score = score
                best_match = i
        # 仅返回超过阈值的匹配
        return (best_match, best_score) if best_score >= threshold else (None, 0)

    # 为issues表创建主题列表用于匹配
    issue_subjects = issues_df['主题'].tolist()

    # 创建合并结果DataFrame
    merged_data = []

    # 使用tqdm添加进度条
    for i, alm_row in tqdm(alm_df.iterrows(), total=len(alm_df), desc="匹配进度"):
        nonconformity = str(alm_row['不符合现象'])
        # 查找最佳匹配
        match_idx, match_score = find_best_match(
            nonconformity, issue_subjects, threshold)

        # 准备合并行数据
        row_data = alm_row.to_dict()
        # 添加匹配信息
        row_data['匹配分数'] = match_score

        # 如果找到匹配，添加issues表数据
        if match_idx is not None:
            matched_issue = issues_df.iloc[match_idx]
            # 修改点：为所有issues列添加前缀（无论列名是否冲突）
            for col in issues_df.columns:
                # 添加"Issues_"前缀
                row_data[f"Issues_{col}"] = matched_issue[col]
            # 添加主题匹配结果（单独列出便于查看）
            row_data['主题匹配结果'] = matched_issue['主题']
            row_data['匹配状态'] = '成功匹配'
        else:
            # 未匹配时仍需添加issues列（保持结构一致）
            for col in issues_df.columns:
                row_data[f"Issues_{col}"] = ""
            row_data['主题匹配结果'] = ""
            row_data['匹配状态'] = '未找到匹配'

        merged_data.append(row_data)

    # 创建合并后的DataFrame
    merged_df = pd.DataFrame(merged_data)

    # 调整列顺序以便阅读（保持原有逻辑）
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

    return merged_df


# 使用示例
if __name__ == "__main__":
    alm_file = r'D:\20220916\laiye\东风项目\流程录屏和相关资料\QIS，redmine表格汇总自动化\ALM任务导出_1755238307929.csv'
    issues_file = r'D:\20220916\laiye\东风项目\流程录屏和相关资料\QIS，redmine表格汇总自动化\issues (1).csv'
    output_file = r'DFcode\alm_issues合并结果.csv'

    result = merge_alm_issues(
        alm_path=alm_file,
        issues_path=issues_file,
        output_path=output_file,
        alm_encoding='ANSI',
        issues_encoding='ANSI',
        threshold=75
    )
