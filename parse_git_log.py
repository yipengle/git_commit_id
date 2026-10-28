# 文件名：parse_git_log.py
import subprocess
import re
import pandas as pd

def parse_git_log():
    print("正在通过 repo forall 抓取各个子仓库的日志...")
    
    # 使用 repo forall 遍历所有子项目执行 git log
    # -p 打印当前项目路径；-c 后跟要执行的命令
    # 这里的命令提取：哈希|作者|日期|标题|正文
    cmd = "repo forall -p -c 'git log -n 50 --pretty=format:\"%H|%an|%ad|%s|%b\" --date=iso'"
    
    try:
        log_text = subprocess.check_output(cmd, shell=True, stderr=subprocess.STDOUT, text=True, errors='ignore')
    except subprocess.CalledProcessError as e:
        print(f"获取 git log 失败: {e.output}")
        return

    data_list = []
    current_project = "未知仓库"

    # 按行解析输出结果
    for line in log_text.strip().split('\n'):
        line = line.strip()
        if not line:
            continue
            
        # 捕获项目路径 (repo forall -p 输出的格式通常是 "project path/to/repo/")
        if line.startswith('project '):
            current_project = line.replace('project ', '').strip()
            continue
            
        # 解析 git log 的内容
        parts = line.split('|', 4)
        if len(parts) >= 4:
            commit_hash = parts[0]
            author = parts[1]
            date = parts[2]
            subject = parts[3]
            body = parts[4] if len(parts) > 4 else ""

            record = {
                'Commit ID': commit_hash,
                'Author': author,
                'Date': date,
                'Project Path': current_project,
                'Subject': subject,
            }

            # 正则匹配你截图中的特定 Tag，兼容中英文冒号
            bug_id_match = re.search(r'\[bug_id\]\s*[:：]?\s*(.*)', body)
            if bug_id_match: record['bug_id'] = bug_id_match.group(1).strip()

            project_tag_match = re.search(r'\[project\]\s*[:：]?\s*(.*)', body)
            if project_tag_match: record['project'] = project_tag_match.group(1).strip()

            desc_match = re.search(r'\[问题描述\]\s*[:：]?\s*(.*)', body)
            if desc_match: record['问题描述'] = desc_match.group(1).strip()

            common_match = re.search(r'\[共性问题\]\s*[:：]?\s*(.*)', body)
            if common_match: record['共性问题'] = common_match.group(1).strip()

            change_id_match = re.search(r'Change-Id:\s*(.*)', body)
            if change_id_match: record['Change-Id'] = change_id_match.group(1).strip()

            # 如果这个 commit 包含我们要的任何标签，就保存下来
            if 'bug_id' in record or '问题描述' in record:
                data_list.append(record)

    # 生成 Excel
    if data_list:
        df = pd.DataFrame(data_list)
        # 整理列顺序
        cols = ['Commit ID', 'Author', 'Date', 'Project Path', 'bug_id', 'project', '问题描述', '共性问题', 'Change-Id', 'Subject']
        exist_cols = [c for c in cols if c in df.columns]
        df = df[exist_cols]

        output_file = 'git_commit_log.xlsx'
        df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"✅ 成功导出 {len(data_list)} 条记录到 {output_file}")
    else:
        print("⚠️ 未解析到任何符合格式的提交记录。")

if __name__ == "__main__":
    parse_git_log()
