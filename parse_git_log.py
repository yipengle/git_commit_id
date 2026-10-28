# 文件名：parse_git_log.py
import subprocess
import re
import pandas as pd

def parse_git_log():
    print("正在抓取当前目录下的 git log...")
    
    # 直接执行 git log，获取最近 100 条
    cmd = ['git', 'log', '-n', '100']
    try:
        # 运行命令并获取输出
        log_text = subprocess.check_output(cmd, text=True, errors='ignore')
    except subprocess.CalledProcessError as e:
        print(f"获取 git log 失败，请确认当前是否在 git 仓库目录下: {e}")
        return

    # 根据 "commit " 关键字切分每一条日志
    # 比如切成 ["commit 59cc...", "commit efd3...", "commit 4b45..."]
    commits = log_text.split('\ncommit ')
    
    data_list = []

    for idx, block in enumerate(commits):
        if not block.strip():
            continue
            
        # 补上第一个被切掉的 "commit " 前缀
        if idx > 0:
            block = 'commit ' + block
            
        record = {}
        
        # 1. 提取 Commit ID (截取前 40 位哈希值)
        commit_id_match = re.search(r'^commit ([a-f0-9]{40})', block)
        if commit_id_match:
            record['Commit ID'] = commit_id_match.group(1)
            
        # 2. 提取 Author
        author_match = re.search(r'Author:\s+([^<]+)', block)
        if author_match:
            record['Author'] = author_match.group(1).strip()
            
        # 3. 提取 Date
        date_match = re.search(r'Date:\s+(.*)', block)
        if date_match:
            record['Date'] = date_match.group(1).strip()
            
        # 4. 提取具体的业务标签（兼容全角/半角冒号，支持有/无空格）
        # [bug_id]
        bug_id_match = re.search(r'\[bug_id\]\s*[:：]?\s*(.*)', block)
        if bug_id_match: record['bug_id'] = bug_id_match.group(1).strip()
        
        # [project]
        proj_match = re.search(r'\[project\]\s*[:：]?\s*(.*)', block)
        if proj_match: record['project'] = proj_match.group(1).strip()
        
        # [问题描述]
        desc_match = re.search(r'\[问题描述\]\s*[:：]?\s*(.*)', block)
        if desc_match: record['问题描述'] = desc_match.group(1).strip()
        
        # [共性问题]
        common_match = re.search(r'\[共性问题\]\s*[:：]?\s*(.*)', block)
        if common_match: record['共性问题'] = common_match.group(1).strip()
        
        # Change-Id
        change_id_match = re.search(r'Change-Id:\s*(.*)', block)
        if change_id_match: record['Change-Id'] = change_id_match.group(1).strip()

        # 如果提取到了数据，加入列表
        if record:
            data_list.append(record)

    # 5. 导出为 Excel
    if data_list:
        df = pd.DataFrame(data_list)
        # 整理列的顺序
        cols = ['Commit ID', 'Author', 'Date', 'bug_id', 'project', '问题描述', '共性问题', 'Change-Id']
        # 过滤掉不存在的列
        exist_cols = [c for c in cols if c in df.columns]
        df = df[exist_cols]

        # 注意：脚本在 fibocom-mtk 目录下执行，我们要把 Excel 生成到上一级目录（Jenkins 工作空间根目录），方便归档
        output_file = '../git_commit_log.xlsx'
        df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"✅ 成功导出 {len(data_list)} 条记录到 {output_file}")
    else:
        print("⚠️ 未解析到任何符合格式的提交记录。")

if __name__ == "__main__":
    parse_git_log()
