import pandas as pd
import re
import html
import os
from datetime import datetime


def _build_pattern_from_query(query):
    # Build a regex alternation pattern from a search query. Handles * wildcard and removes common boolean operators.
    if not query or not isinstance(query, str):
        return None
    tokens = re.split(r"\s+", query)
    cleaned = []
    skip_next = False  # 追踪 NOT 操作符

    for t in tokens:
        up = t.upper()

        # 遇到 NOT 操作符，标记跳过下一个词
        if up == "NOT":
            skip_next = True
            continue

        # 跳过 AND、OR 并重置 NOT 标志
        if up in ("AND", "OR"):
            skip_next = False
            continue

        # 移除括号和字段限定符
        t = t.strip('()')
        if '[' in t:
            t = t.split('[')[0]
        t = t.strip('"')

        # 忽略无效的 token
        if not re.search(r"[A-Za-z0-9*]", t):
            continue

        # 跳过 NOT 操作的词（如： NOT cancer）
        if skip_next:
            skip_next = False
            continue

        cleaned.append(t)

    patterns = []
    for t in cleaned:
        if '*' in t:
            # 修改：使用占位符，避免被 replace() 误伤
            esc = ''.join([re.escape(ch) if ch != '*' else '___WILDCARD___' for ch in t])
            tmp = esc.replace('___WILDCARD___', r'\w*')
        else:
            tmp = re.escape(t)
        patterns.append(tmp)

    if not patterns:
        return None
    return r'(?i)(' + '|'.join(patterns) + r')'


def generate_reading_list(input_path_or_df, output_html_path, search_info=None):
    # Generate a night-mode HTML reading list from CSV/Excel or a DataFrame with interactive features.
    # Optional search_info dict may contain 'search_keywords', 'paper_type', 'release_date_cutoff', 'grab_total', 'save_path', 'search_date'.
    try:
        if isinstance(input_path_or_df, pd.DataFrame):
            df = input_path_or_df
        else:
            input_path = str(input_path_or_df)
            _, ext = os.path.splitext(input_path)
            ext = ext.lower()
            if ext in ('.xls', '.xlsx'):
                df = pd.read_excel(input_path, sheet_name=0)
            else:
                df = pd.read_csv(input_path)
    except Exception as e:
        print(f"Failed to read input: {e}")
        return

    pattern = None
    if search_info and 'search_keywords' in search_info:
        pattern = _build_pattern_from_query(search_info.get('search_keywords'))
    if not pattern:
        sample = ''
        if 'Title' in df.columns or 'TI' in df.columns:
            col = 'Title' if 'Title' in df.columns else 'TI'
            vals = df[col].dropna().values
            sample = str(vals[0]) if len(vals)>0 else ''
            words = re.findall(r"[A-Za-z0-9]{3,}", sample)
            if words:
                pattern = r'(?i)(' + re.escape(words[0]) + r')'

    colors = ['#ffd54f', '#ff79c6', '#8be9fd', '#50fa7b', '#ffb86b']

    def _make_highlighter(pat):
        if not pat:
            return lambda s: s
        prog = re.compile(pat)
        counter = {'i': 0}
        def repl(m):
            idx = counter['i'] % len(colors)
            color = colors[idx]
            counter['i'] += 1
            return f'<span style="color: {color}; font-weight:700;">{m.group(0)}</span>'
        return lambda s: prog.sub(repl, s)

    highlighter = _make_highlighter(pattern)

    def truncate_text(text, length=1500):
        if not isinstance(text, str):
            return ""
        if len(text) > length:
            return text[:length] + "..."
        return text

    search_block_html = ''
    if search_info:
        sd = search_info.get('search_date') or datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        sk = search_info.get('search_keywords', 'N/A')
        pt = search_info.get('paper_type', 'N/A')
        rc = search_info.get('release_date_cutoff', None)
        rc_text = f"last {rc} days" if rc else 'all time'
        gt = search_info.get('grab_total_requested', search_info.get('grab_total', 'all'))
        savep = search_info.get('save_path', '')

        # Search summary block displayed on top of the HTML
        search_block_html = f'''\
        <div class="search-summary" id="search-summary">\
            <h1>Search Summary (Night mode)</h1>\
            <div class="search-meta">\
                <div><strong>Search time:</strong> {sd}</div>\
                <div><strong>Query:</strong> <code class="query">{html.escape(sk)}</code></div>\
                <div><strong>Paper type:</strong> {html.escape(str(pt))}  <strong>Time range:</strong> {html.escape(rc_text)}</div>\
                <div><strong>Requested count:</strong> {html.escape(str(gt))}  <strong>Save path:</strong> {html.escape(str(savep))}</div>\
            </div>\
        </div>\
        '''

    # 添加交互式JavaScript和CSS

    # Generate sidebar bookmark links
    sidebar_links_html = ""
    for idx, row in df.iterrows():
        # 使用实际的Excel列名
        journal_raw = row.get('Journal', row.get('Journal (TA)', row.get('TA', '')))
        journal = str(journal_raw).strip() if pd.notna(journal_raw) else "Unknown"
        
        pub_date_raw = row.get('publish_date', row.get('Publish Date (LR)', row.get('LR', '')))
        if pd.notna(pub_date_raw) and str(pub_date_raw).strip():
            pub_date = str(pub_date_raw).replace("-", "").replace("/", "").replace(" ", "")
        else:
            pub_date = "Unknown"
        bookmark_text = f"{journal}. {pub_date}"
        # 添加状态指示器容器
        sidebar_links_html += f'            <li><a href="#article-{idx}" data-article-id="{idx}"><span class="bookmark-indicators" id="indicators-{idx}"></span>{html.escape(bookmark_text)}</a></li>\n'

    html_content = f'''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Reading List (Night mode)</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
            /* Sidebar styles */
            .sidebar {{ position: fixed; left: 0; top: 0; width: 280px; height: 100%; background: #2a2a2a; border-right: 1px solid #444; overflow-y: auto; padding: 20px; z-index: 1000; transition: transform 0.3s; }}
            .sidebar.hidden {{ transform: translateX(-280px); }}
            .sidebar h2 {{ font-size: 18px; margin-bottom: 15px; color: #4a9eff; }}
            .sidebar ul {{ list-style: none; }}
            .sidebar li {{ margin: 8px 0; }}
            .sidebar a {{ color: #b0b0b0; text-decoration: none; font-size: 14px; display: flex; align-items: center; gap: 5px; padding: 5px; border-radius: 3px; transition: all 0.2s; }}
            .sidebar a:hover {{ background: #3a3a3a; color: #4a9eff; }}
            .sidebar-toggle {{ position: fixed; left: 290px; top: 20px; background: #4a9eff; color: white; border: none; padding: 10px 15px; cursor: pointer; border-radius: 5px; z-index: 999; transition: left 0.3s; }}
            .sidebar-toggle.sidebar-hidden {{ left: 10px; }}
            
            :root {{ --bg:#0b1220; --card:#07101a; --muted:#98a2b3; --text:#e6eef8; --accent:#66d9ef; --metric-bg:rgba(255,255,255,0.04); --border:rgba(255,255,255,0.06); }}
            html,body {{ background: linear-gradient(180deg,#051021 0%,#071827 100%); color:var(--text); font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif; margin:0; padding:0; transition: padding-left 0.3s; }}
            body {{ padding-left: 300px; }}
            body.sidebar-closed {{ padding-left: 0; }}
            .bookmark-indicators {{ display: inline-flex; gap: 3px; min-width: 35px; flex-shrink: 0; }}
            .bookmark-indicators .indicator {{ font-size: 14px; }}
            .bookmark-indicators .star-indicator {{ color: #ffd700; }}
            .bookmark-indicators .read-indicator {{ color: #4CAF50; }}
            .container {{ max-width:1000px; margin:24px auto; padding:18px }}
            .search-summary {{ background: linear-gradient(90deg, rgba(255,255,255,0.02), rgba(255,255,255,0.01)); border:1px solid var(--border); padding:20px; border-radius:12px; margin-bottom:20px; }}
            .search-summary h1 {{ margin:0 0 8px 0; color:var(--accent); font-size:1.6em }}
            .search-meta div {{ margin:6px 0; color:var(--muted) }}
            .query {{ background: rgba(255,255,255,0.03); padding:6px 8px; border-radius:6px; color:var(--text); font-family:monospace }}

            .article-card {{ background:var(--card); padding:30px; margin-bottom:18px; box-shadow: 0 6px 18px rgba(2,6,23,0.6); border:1px solid var(--border); border-radius:10px; page-break-inside:avoid; position:relative; transition: border-color 0.3s }}
            .article-card.starred {{ border-left: 4px solid #ffd700; }}
            .article-card.read {{ opacity: 0.6; }}
            
            .article-title {{ color:var(--accent); font-size:1.3em; font-weight:700; margin-bottom:8px }}
            .article-meta {{ color:var(--muted); font-size:0.95em; margin-bottom:14px }}
            .journal-info {{ font-style:italic; color:var(--text); font-weight:600 }}
            .metrics {{ display:inline-block; background:var(--metric-bg); padding:4px 8px; border-radius:6px; margin-right:6px; color:var(--text); font-size:0.85em }}
            .abstract-section {{ margin-top:12px }}
            .abstract-label {{ font-weight:700; color:var(--text); margin-bottom:6px; display:block }}
            .abstract-text {{ color:#dbe9f6; line-height:1.7; text-align:justify }}
            .article-ids {{ margin-top:16px; color:var(--muted); font-size:0.9em; border-top:1px dashed rgba(255,255,255,0.03); padding-top:10px }}
            
            /* 交互按钮样式 */
            .action-buttons {{ position: absolute; top: 20px; right: 20px; display: flex; gap: 8px; }}
            .action-btn {{ cursor: pointer; padding: 6px 10px; border-radius: 6px; font-size: 0.9em; transition: all 0.2s; background: rgba(255,255,255,0.05); border: 1px solid var(--border); color: var(--muted); }}
            .action-btn:hover {{ background: rgba(255,255,255,0.1); transform: scale(1.05); }}
            .action-btn.active {{ background: rgba(102,217,239,0.2); color: #66d9ef; border-color: #66d9ef; }}
            .bookmark-btn.active {{ color: #ffd700; border-color: #ffd700; }}
            .star-btn.active {{ color: #ffd700; border-color: #ffd700; }}
            .read-btn.active {{ color: #50fa7b; border-color: #50fa7b; }}
            
            @media print {{ body{{ background:white; color:black }} .article-card{{ box-shadow:none; border:none }} .action-buttons {{ display: none; }} }}
        </style>
    </head>
    <body>
    <button class="sidebar-toggle" onclick="toggleSidebar()">☰</button>
    <div class="sidebar">
        <h2>📑 Bookmarks</h2>
        <ul>
            <li><a href="#search-summary">Research Summary</a></li>
    {sidebar_links_html}
        </ul>
    </div>
    <div class="container">
    {search_block_html}
    '''

    for index, row in df.iterrows():
        title = str(row.get('Title', row.get('TI', 'No Title')))
        journal = str(row.get('Journal', row.get('TA', '')))
        publish_date = str(row.get('publish_date', row.get('LR', '')))
        abstract = str(row.get('Abstract', row.get('AB', '')))
        pmid = str(row.get('PMID', ''))
        doi = str(row.get('DOI', row.get('LID', '')))
        impact_factor = str(row.get('IF', ''))
        quartile = str(row.get('JCR_Quartile', row.get('Quartile', '')))

        display_abstract = truncate_text(abstract, length=2000)
        safe_title = html.escape(title)
        safe_abstract = html.escape(display_abstract)
        highlighted_title = highlighter(safe_title) if pattern else safe_title
        highlighted_abstract = highlighter(safe_abstract) if pattern else safe_abstract

        # 创建书签标题（期刊名+日期）
        bookmark_title = f"{journal} - {publish_date}"
        article_id = f"article-{index}"

        meta_html = f'<span class="journal-info">{journal}</span>. {publish_date}.'
        metrics_html = ''
        if impact_factor and impact_factor != 'nan':
            metrics_html += f'<span class="metrics">IF: {impact_factor}</span>'
        if quartile and quartile != 'nan':
            metrics_html += f'<span class="metrics">{quartile}</span>'

        article_html = f'''
        <div class="article-card" id="{article_id}" data-bookmark-title="{html.escape(bookmark_title)}">
            <div class="action-buttons">
                <button class="action-btn star-btn" onclick="toggleStar(this)" title="星标重点">⭐</button>
                <button class="action-btn read-btn" onclick="toggleRead(this)" title="标记已读">✓</button>
            </div>
            <div class="article-title">{highlighted_title}</div>
            <div class="article-meta">
                {meta_html} <br>
                {metrics_html}
            </div>
            <div class="abstract-section">
                <span class="abstract-label">Abstract</span>
                <div class="abstract-text">
                    {highlighted_abstract}
                </div>
            </div>
            <div class="article-ids">
                PMID: {pmid} &nbsp;|&nbsp; DOI: {doi}
            </div>
        </div>
        '''
        html_content += article_html

# 添加交互式JavaScript
    html_content += '''
    <script>
        // 更新侧边栏的小图标
        function updateSidebarIndicator(articleId) {
            // 从 article-0 提取 0
            const articleNum = articleId.replace('article-', ''); 
            const indicatorContainer = document.getElementById('indicators-' + articleNum);
            if (!indicatorContainer) return;
            
            const starred = JSON.parse(localStorage.getItem('starred') || '[]');
            const read = JSON.parse(localStorage.getItem('read') || '[]');
            
            let html = '';
            if (starred.includes(articleId)) {
                html += '<span class="indicator star-indicator">⭐</span>';
            }
            if (read.includes(articleId)) {
                html += '<span class="indicator read-indicator">✓</span>';
            }
            indicatorContainer.innerHTML = html;
        }
        
        // 批量更新所有侧边栏图标
        function updateAllSidebarIndicators() {
            const allLinks = document.querySelectorAll('.sidebar a[data-article-id]');
            allLinks.forEach(link => {
                const articleNum = link.getAttribute('data-article-id'); // 这里拿到的是数字索引
                if (articleNum !== null) {
                    const articleId = 'article-' + articleNum;
                    updateSidebarIndicator(articleId);
                }
            });
        }
        
        function toggleSidebar() {
            const sidebar = document.querySelector('.sidebar');
            const toggle = document.querySelector('.sidebar-toggle');
            const body = document.body;
            
            sidebar.classList.toggle('hidden');
            toggle.classList.toggle('sidebar-hidden');
            body.classList.toggle('sidebar-closed');
        }
        
        // --- 修复重点：类名修正为 .article-card，并增加按钮状态切换 ---
        
        function toggleStar(btn) {
            // 修复1：这里必须用 .article-card，因为你的HTML生成的class是这个
            const card = btn.closest('.article-card'); 
            if (!card) return;
            const articleId = card.id;
            
            let starred = JSON.parse(localStorage.getItem('starred') || '[]');
            const index = starred.indexOf(articleId);
            
            if (index > -1) {
                starred.splice(index, 1);
                card.classList.remove('starred');
                btn.classList.remove('active'); // 修复2：同步移除按钮高亮
            } else {
                starred.push(articleId);
                card.classList.add('starred');
                btn.classList.add('active'); // 修复2：同步添加按钮高亮
            }
            
            localStorage.setItem('starred', JSON.stringify(starred));
            updateSidebarIndicator(articleId); // 更新侧边栏
        }
        
        function toggleRead(btn) {
            // 修复1：同上，修正类名
            const card = btn.closest('.article-card');
            if (!card) return;
            const articleId = card.id;
            
            let read = JSON.parse(localStorage.getItem('read') || '[]');
            const index = read.indexOf(articleId);
            
            if (index > -1) {
                read.splice(index, 1);
                card.classList.remove('read');
                btn.classList.remove('active'); // 修复2
            } else {
                read.push(articleId);
                card.classList.add('read');
                btn.classList.add('active'); // 修复2
            }
            
            localStorage.setItem('read', JSON.stringify(read));
            updateSidebarIndicator(articleId); // 更新侧边栏
        }
        
        window.onload = function() {
            const starred = JSON.parse(localStorage.getItem('starred') || '[]');
            const read = JSON.parse(localStorage.getItem('read') || '[]');
            
            // 恢复星标状态
            starred.forEach(id => {
                const card = document.getElementById(id);
                if (card) {
                    card.classList.add('starred');
                    // 修复3：加载时也点亮按钮
                    const btn = card.querySelector('.star-btn');
                    if (btn) btn.classList.add('active');
                }
            });
            
            // 恢复已读状态
            read.forEach(id => {
                const card = document.getElementById(id);
                if (card) {
                    card.classList.add('read');
                    // 修复3：加载时也点亮按钮
                    const btn = card.querySelector('.read-btn');
                    if (btn) btn.classList.add('active');
                }
            });
            
            // 初始化侧边栏所有图标
            updateAllSidebarIndicators();
        }
    </script>
    </div>
    </body>
    </html>
    '''

    out_dir = os.path.dirname(output_html_path) or '.'
    os.makedirs(out_dir, exist_ok=True)
    with open(output_html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"Conversion complete: {output_html_path}")


if __name__ == "__main__":
    input_csv = "wnt5a_fibro.xlsx - Sheet.csv"
    output_html = "reading_list.html"
    generate_reading_list(input_csv, output_html)
    html_parts.append('            .sidebar { position: fixed; left: 0; top: 0; width: 280px; height: 100%; background: #2a2a2a; border-right: 1px solid #444; overflow-y: auto; padding: 20px; z-index: 1000; transition: transform 0.3s; }\\n')
    html_parts.append('            .sidebar.hidden { transform: translateX(-280px); }\\n')
    html_parts.append('            .sidebar h2 { font-size: 18px; margin-bottom: 15px; color: #4a9eff; }\\n')
    html_parts.append('            .sidebar ul { list-style: none; }\\n')
    html_parts.append('            .sidebar li { margin: 8px 0; }\\n')
    html_parts.append('            .sidebar a { color: #b0b0b0; text-decoration: none; font-size: 14px; display: block; padding: 5px; border-radius: 3px; transition: all 0.2s; }\\n')
    html_parts.append('            .sidebar a:hover { background: #3a3a3a; color: #4a9eff; }\\n')
    html_parts.append('            .sidebar-toggle { position: fixed; left: 290px; top: 20px; background: #4a9eff; color: white; border: none; padding: 10px 15px; cursor: pointer; border-radius: 5px; z-index: 999; transition: left 0.3s; }\\n')
    html_parts.append('            .sidebar-toggle.sidebar-hidden { left: 10px; }\\n')
