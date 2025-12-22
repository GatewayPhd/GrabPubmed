import os
print("working directory:", os.getcwd())
from Bio import Entrez, Medline
import requests
import openpyxl
import time
from tqdm import trange
from bs4 import BeautifulSoup
import re

class pubmed_utils():
    def __init__(self):
        pass
        
        
    def get_main_info_into_excel(self, api_key, search_key_words, release_date_cutoff=None, paper_type="Article", grab_total=None, save_path="./paper_info.xlsx"):
        '''
        grab info from pubmed using NCBI eUtils API, save it into a excel
        支持逻辑符号: AND, OR, NOT 等
        
        Parameters:
        -----------
        api_key : str
            NCBI eUtils API key
        search_key_words : str
            Search query (支持逻辑符号，如 "wnt5a AND cancer")
        release_date_cutoff : int, optional
            发布时间范围（天数），默认为None（所有数据）
        paper_type : str, optional
            论文类型，默认为"Article"
        grab_total : int, optional
            获取论文数量，默认为None（获取所有）
        save_path : str
            Excel保存路径
        '''
        
        grab_step = 10
        
        # 构建搜索词
        search_term = search_key_words
        if paper_type:
            search_term += f" AND \"{paper_type}\"[PT]"
        
        # 步骤1: ESearch - 搜索论文
        esearch_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
        esearch_params = {
            "db": "pubmed",
            "term": search_term,
            "api_key": api_key,
            "usehistory": "y",
            "retmax": 0  # 只获取总数
        }
        
        # 添加日期范围限制
        if release_date_cutoff:
            esearch_params["reldate"] = release_date_cutoff
        
        print("Searching PubMed...")
        esearch_response = requests.get(esearch_url, params=esearch_params)
        esearch_data = esearch_response.text
        
        # 解析搜索结果
        import xml.etree.ElementTree as ET
        root = ET.fromstring(esearch_data)
        total = int(root.find("Count").text)
        webenv = root.find("WebEnv").text
        query_key = root.find("QueryKey").text
        
        print(f"Find total: {total}")
        
        if grab_total is None or grab_total > total:
            grab_total = total
        
        # 初始化Excel
        self.excel_property_dic = {token:index for index, token in enumerate(["PMID", "TI", "TA", "IF", "Quartile", "JCR_Quartile", "Top", "OA", "LR", "AB", "LID"], start=1)}
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=self.excel_property_dic["PMID"]).value = "PMID"
        ws.cell(row=1, column=self.excel_property_dic["TI"]).value = "Title"
        ws.cell(row=1, column=self.excel_property_dic["TA"]).value = "Journal"
        ws.cell(row=1, column=self.excel_property_dic["IF"]).value = "IF"
        ws.cell(row=1, column=self.excel_property_dic["Quartile"]).value = "JCR_Quartile"
        ws.cell(row=1, column=self.excel_property_dic["JCR_Quartile"]).value = "CSA_Quartile"
        ws.cell(row=1, column=self.excel_property_dic["Top"]).value = "Top"
        ws.cell(row=1, column=self.excel_property_dic["OA"]).value = "Open Access"
        ws.cell(row=1, column=self.excel_property_dic["LR"]).value = "publish_date"
        ws.cell(row=1, column=self.excel_property_dic["AB"]).value = "Abstract"
        ws.cell(row=1, column=self.excel_property_dic["LID"]).value = "DOI"

        # 步骤2: EFetch - 获取详细信息
        cur_row = 2
        efetch_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
        
        for step in trange(0, (grab_total + grab_step - 1) // grab_step, desc="getting pubmed info"):
            efetch_params = {
                "db": "pubmed",
                "retstart": step * grab_step,
                "retmax": grab_step,
                "webenv": webenv,
                "query_key": query_key,
                "rettype": "medline",
                "retmode": "text",
                "api_key": api_key
            }
            
            efetch_response = requests.get(efetch_url, params=efetch_params)
            response_text = efetch_response.text
            
            # 修复：使用正则表达式按照 PMID 行来分割记录
            # PMID行格式为: "PMID- 12345678"
            record_texts = re.split(r'\n(?=PMID- )', response_text)
            
            for record_text in record_texts:
                if not record_text.strip() or not record_text.startswith('PMID-'):
                    continue
                
                try:
                    # 解析单条记录
                    records = list(Medline.parse(record_text.split('\n')))
                    
                    for record in records:
                        if 'PMID' not in record:
                            continue
                            
                        # 写入Excel - 每个字段
                        for key in self.excel_property_dic.keys():
                            if key not in record:
                                continue
                            
                            key_info = record[key]
                            
                            # 处理列表类型的字段
                            if isinstance(key_info, list):
                                if key == 'LID':  # DOI 字段
                                    # 找到包含 [doi] 的项
                                    doi_items = [item for item in key_info if '[doi]' in item.lower()]
                                    if doi_items:
                                        key_info = doi_items[0]
                                    elif key_info:
                                        key_info = key_info[0]
                                    else:
                                        key_info = ''
                                elif key == 'LR':  # 日期字段 - 取最新的（第一个）
                                    key_info = key_info[0] if key_info else ''
                                else:
                                    # 其他列表字段不应该出现在这些关键字段中
                                    # 如果出现，用分号连接
                                    key_info = '; '.join(str(x) for x in key_info)
                            
                            ws.cell(row=cur_row, column=self.excel_property_dic[key]).value = key_info
                        
                        cur_row += 1
                        
                except Exception as e:
                    print(f"解析记录时出错: {e}")
                    continue
            
            time.sleep(0.5)  # 遵守API限制

        wb.save(save_path)
        print(f"Data saved to {save_path}")
        print(f"Total records written: {cur_row - 2}")
        
        
    def embed_IF_into_excel(self, excel_path, jcr_csa_path="JCR_CSA_2025.xlsx"):
        '''
        grab IF, JCR Quartile, CSA Quartile, Top, and OA info from local JCR_CSA_2025.xlsx and save it into excel
        支持全称和缩略名双重匹配
        '''
        
        # Load JCR_CSA data
        jcr_csa_wb = openpyxl.load_workbook(jcr_csa_path)
        
        # Sheet 1: JCR info (IF and Quartile)
        ws_jcr = jcr_csa_wb[jcr_csa_wb.sheetnames[0]]
        jcr_dic = {}  # {journal_name: {"IF": value, "Quartile": value}}
        jcr_abbr_dic = {}  # {abbreviation: {"IF": value, "Quartile": value}}
        
        for jcr_row in range(2, ws_jcr.max_row+1):
            journal_name = ws_jcr.cell(row=jcr_row, column=1).value  # 期刊全称
            journal_abbr = ws_jcr.cell(row=jcr_row, column=2).value  # 期刊缩略名
            jif_value = ws_jcr.cell(row=jcr_row, column=7).value     # 2024JIF (列位置已调整)
            quartile_value = ws_jcr.cell(row=jcr_row, column=8).value  # Quartile (列位置已调整)
            
            if journal_name:
                jcr_dic[journal_name.strip().upper()] = {"IF": jif_value, "Quartile": quartile_value}
            if journal_abbr:
                jcr_abbr_dic[journal_abbr.strip().upper()] = {"IF": jif_value, "Quartile": quartile_value}
        
        # Sheet 2: CSA info (中科院分区, Top, OA)
        ws_csa = jcr_csa_wb[jcr_csa_wb.sheetnames[1]]
        csa_dic = {}  # {journal_name: {"CSA_Quartile": value, "Top": value, "OA": value}}
        csa_abbr_dic = {}  # {abbreviation: {"CSA_Quartile": value, "Top": value, "OA": value}}
        
        for csa_row in range(2, ws_csa.max_row+1):
            journal_name = ws_csa.cell(row=csa_row, column=1).value  # 期刊全称
            journal_abbr = ws_csa.cell(row=csa_row, column=2).value  # 期刊缩略名
            csa_quartile = ws_csa.cell(row=csa_row, column=3).value   # 2025分区 (列位置已调整)
            top_value = ws_csa.cell(row=csa_row, column=4).value      # Top (列位置已调整)
            oa_value = ws_csa.cell(row=csa_row, column=5).value       # Open Access (列位置已调整)
            
            if journal_name:
                csa_dic[journal_name.strip().upper()] = {"CSA_Quartile": csa_quartile, "Top": top_value, "OA": oa_value}
            if journal_abbr:
                csa_abbr_dic[journal_abbr.strip().upper()] = {"CSA_Quartile": csa_quartile, "Top": top_value, "OA": oa_value}
        
        # Load target excel and update values
        wb = openpyxl.load_workbook(excel_path)
        ws = wb["Sheet"]
        
        # 匹配统计
        match_stats = {
            "jcr_full_match": 0,
            "jcr_abbr_match": 0,
            "jcr_partial_match": 0,
            "jcr_no_match": 0,
            "csa_full_match": 0,
            "csa_abbr_match": 0,
            "csa_partial_match": 0,
            "csa_no_match": 0
        }
        fail_list = []
        
        for cur_row in range(2, ws.max_row+1):
            j_name = ws.cell(row=cur_row, column=self.excel_property_dic["TA"]).value
            if not j_name:
                continue
                
            j_name_upper = j_name.strip().upper()
            jcr_found = False
            csa_found = False
            match_method = ""
            
            # === JCR 匹配策略 ===
            # 1. 尝试全称精确匹配
            if j_name_upper in jcr_dic:
                ws.cell(row=cur_row, column=self.excel_property_dic["IF"]).value = jcr_dic[j_name_upper]["IF"]
                ws.cell(row=cur_row, column=self.excel_property_dic["Quartile"]).value = jcr_dic[j_name_upper]["Quartile"]
                jcr_found = True
                match_stats["jcr_full_match"] += 1
                match_method = "JCR:全称"
            
            # 2. 尝试缩略名精确匹配
            elif j_name_upper in jcr_abbr_dic:
                ws.cell(row=cur_row, column=self.excel_property_dic["IF"]).value = jcr_abbr_dic[j_name_upper]["IF"]
                ws.cell(row=cur_row, column=self.excel_property_dic["Quartile"]).value = jcr_abbr_dic[j_name_upper]["Quartile"]
                jcr_found = True
                match_stats["jcr_abbr_match"] += 1
                match_method = "JCR:缩略"
            
            # 3. 尝试部分匹配（全称）
            else:
                for jcr_journal in jcr_dic.keys():
                    if j_name_upper in jcr_journal or jcr_journal in j_name_upper:
                        ws.cell(row=cur_row, column=self.excel_property_dic["IF"]).value = jcr_dic[jcr_journal]["IF"]
                        ws.cell(row=cur_row, column=self.excel_property_dic["Quartile"]).value = jcr_dic[jcr_journal]["Quartile"]
                        jcr_found = True
                        match_stats["jcr_partial_match"] += 1
                        match_method = "JCR:部分"
                        break
            
            # 4. 未匹配
            if not jcr_found:
                ws.cell(row=cur_row, column=self.excel_property_dic["IF"]).value = "Unknow"
                ws.cell(row=cur_row, column=self.excel_property_dic["Quartile"]).value = "Unknow"
                match_stats["jcr_no_match"] += 1
                match_method = "JCR:未匹配"
            
            # === CSA 匹配策略 ===
            # 1. 尝试全称精确匹配
            if j_name_upper in csa_dic:
                ws.cell(row=cur_row, column=self.excel_property_dic["JCR_Quartile"]).value = csa_dic[j_name_upper]["CSA_Quartile"]
                ws.cell(row=cur_row, column=self.excel_property_dic["Top"]).value = csa_dic[j_name_upper]["Top"]
                ws.cell(row=cur_row, column=self.excel_property_dic["OA"]).value = csa_dic[j_name_upper]["OA"]
                csa_found = True
                match_stats["csa_full_match"] += 1
                match_method += " | CSA:全称"
            
            # 2. 尝试缩略名精确匹配
            elif j_name_upper in csa_abbr_dic:
                ws.cell(row=cur_row, column=self.excel_property_dic["JCR_Quartile"]).value = csa_abbr_dic[j_name_upper]["CSA_Quartile"]
                ws.cell(row=cur_row, column=self.excel_property_dic["Top"]).value = csa_abbr_dic[j_name_upper]["Top"]
                ws.cell(row=cur_row, column=self.excel_property_dic["OA"]).value = csa_abbr_dic[j_name_upper]["OA"]
                csa_found = True
                match_stats["csa_abbr_match"] += 1
                match_method += " | CSA:缩略"
            
            # 3. 尝试部分匹配（全称）
            else:
                for csa_journal in csa_dic.keys():
                    if j_name_upper in csa_journal or csa_journal in j_name_upper:
                        ws.cell(row=cur_row, column=self.excel_property_dic["JCR_Quartile"]).value = csa_dic[csa_journal]["CSA_Quartile"]
                        ws.cell(row=cur_row, column=self.excel_property_dic["Top"]).value = csa_dic[csa_journal]["Top"]
                        ws.cell(row=cur_row, column=self.excel_property_dic["OA"]).value = csa_dic[csa_journal]["OA"]
                        csa_found = True
                        match_stats["csa_partial_match"] += 1
                        match_method += " | CSA:部分"
                        break
            
            # 4. 未匹配
            if not csa_found:
                match_stats["csa_no_match"] += 1
                match_method += " | CSA:未匹配"
            
            # 记录未完全匹配的期刊
            if not (jcr_found and csa_found):
                fail_list.append(f"{j_name[:40]} ({match_method})")
        
        # 打印详细匹配统计
        total_journals = ws.max_row - 1
        print("\n" + "="*60)
        print("期刊信息匹配报告")
        print("="*60)
        print(f"总期刊数: {total_journals}\n")
        
        print("JCR (IF & Quartile) 匹配情况:")
        print(f"  - 全称精确匹配: {match_stats['jcr_full_match']} ({match_stats['jcr_full_match']/total_journals*100:.1f}%)")
        print(f"  - 缩略名精确匹配: {match_stats['jcr_abbr_match']} ({match_stats['jcr_abbr_match']/total_journals*100:.1f}%)")
        print(f"  - 部分匹配: {match_stats['jcr_partial_match']} ({match_stats['jcr_partial_match']/total_journals*100:.1f}%)")
        print(f"  - 未匹配: {match_stats['jcr_no_match']} ({match_stats['jcr_no_match']/total_journals*100:.1f}%)")
        jcr_total_matched = match_stats['jcr_full_match'] + match_stats['jcr_abbr_match'] + match_stats['jcr_partial_match']
        print(f"  总匹配率: {jcr_total_matched/total_journals*100:.1f}%\n")
        
        print("CSA (分区 & Top & OA) 匹配情况:")
        print(f"  - 全称精确匹配: {match_stats['csa_full_match']} ({match_stats['csa_full_match']/total_journals*100:.1f}%)")
        print(f"  - 缩略名精确匹配: {match_stats['csa_abbr_match']} ({match_stats['csa_abbr_match']/total_journals*100:.1f}%)")
        print(f"  - 部分匹配: {match_stats['csa_partial_match']} ({match_stats['csa_partial_match']/total_journals*100:.1f}%)")
        print(f"  - 未匹配: {match_stats['csa_no_match']} ({match_stats['csa_no_match']/total_journals*100:.1f}%)")
        csa_total_matched = match_stats['csa_full_match'] + match_stats['csa_abbr_match'] + match_stats['csa_partial_match']
        print(f"  总匹配率: {csa_total_matched/total_journals*100:.1f}%")
        
        if fail_list:
            print(f"\n未完全匹配的期刊 (显示前20个):")
            for item in fail_list[:20]:
                print(f"  {item}")
        
        print("="*60)
        wb.save(excel_path)
    
    

    
    def refine_IF_matching(self, excel_path, jcr_csa_path="JCR_CSA_2025.xlsx", min_similarity=0.6):
        '''
        对已保存的 Excel 文件进行补充匹配，使用更智能的模糊匹配策略
        不调用 PubMed API，仅对未匹配的记录进行二次匹配
        
        Parameters:
        -----------
        excel_path : str
            已保存的 Excel 文件路径
        jcr_csa_path : str
            JCR_CSA 数据文件路径
        min_similarity : float
            最小相似度阈值 (0-1)，默认 0.6
        '''
        
        import difflib
        
        print("\n" + "="*70)
        print("开始智能补充匹配")
        print("="*70)
        
        # 加载 JCR_CSA 数据
        jcr_csa_wb = openpyxl.load_workbook(jcr_csa_path)
        
        # 构建完整的期刊数据库（全称 + 缩略名）
        print("\n加载期刊数据库...")
        ws_jcr = jcr_csa_wb[jcr_csa_wb.sheetnames[0]]
        ws_csa = jcr_csa_wb[jcr_csa_wb.sheetnames[1]]
        
        # JCR 数据字典
        jcr_journals = {}  # {期刊名(大写): {"full": 全称, "abbr": 缩略, "IF": IF值, "Quartile": 分区}}
        for row in range(2, ws_jcr.max_row+1):
            full_name = ws_jcr.cell(row=row, column=1).value
            abbr_name = ws_jcr.cell(row=row, column=2).value
            jif_value = ws_jcr.cell(row=row, column=7).value
            quartile = ws_jcr.cell(row=row, column=8).value
            
            if full_name:
                key = full_name.strip().upper()
                jcr_journals[key] = {
                    "full": full_name.strip(),
                    "abbr": abbr_name.strip() if abbr_name else "",
                    "IF": jif_value,
                    "Quartile": quartile
                }
            if abbr_name:
                key_abbr = abbr_name.strip().upper()
                if key_abbr not in jcr_journals:
                    jcr_journals[key_abbr] = {
                        "full": full_name.strip() if full_name else "",
                        "abbr": abbr_name.strip(),
                        "IF": jif_value,
                        "Quartile": quartile
                    }
        
        # CSA 数据字典
        csa_journals = {}  # {期刊名(大写): {"full": 全称, "abbr": 缩略, "CSA": 分区, "Top": Top, "OA": OA}}
        for row in range(2, ws_csa.max_row+1):
            full_name = ws_csa.cell(row=row, column=1).value
            abbr_name = ws_csa.cell(row=row, column=2).value
            csa_quartile = ws_csa.cell(row=row, column=3).value
            top = ws_csa.cell(row=row, column=4).value
            oa = ws_csa.cell(row=row, column=5).value
            
            if full_name:
                key = full_name.strip().upper()
                csa_journals[key] = {
                    "full": full_name.strip(),
                    "abbr": abbr_name.strip() if abbr_name else "",
                    "CSA_Quartile": csa_quartile,
                    "Top": top,
                    "OA": oa
                }
            if abbr_name:
                key_abbr = abbr_name.strip().upper()
                if key_abbr not in csa_journals:
                    csa_journals[key_abbr] = {
                        "full": full_name.strip() if full_name else "",
                        "abbr": abbr_name.strip(),
                        "CSA_Quartile": csa_quartile,
                        "Top": top,
                        "OA": oa
                    }
        
        print(f"已加载 {len(jcr_journals)} 个 JCR 期刊")
        print(f"已加载 {len(csa_journals)} 个 CSA 期刊")
        
        # 加载目标 Excel
        wb = openpyxl.load_workbook(excel_path)
        ws = wb["Sheet"]
        
        # 统计未匹配的记录
        unmatched_rows = []
        for row in range(2, ws.max_row+1):
            if_value = ws.cell(row=row, column=self.excel_property_dic["IF"]).value
            j_name = ws.cell(row=row, column=self.excel_property_dic["TA"]).value
            if if_value == "Unknow" and j_name:
                unmatched_rows.append((row, j_name.strip()))
        
        print(f"\n找到 {len(unmatched_rows)} 个未匹配的期刊记录")
        
        if not unmatched_rows:
            print("所有记录已匹配，无需补充匹配")
            return
        
        # 模糊匹配函数
        def fuzzy_match_score(pubmed_name, jcr_name):
            '''
            计算两个期刊名称的相似度得分
            策略：
            1. 单词级别的匹配
            2. 前缀匹配
            3. 整体字符串相似度
            '''
            pubmed_upper = pubmed_name.upper()
            jcr_upper = jcr_name.upper()
            
            # 分割成单词
            pubmed_words = [w for w in pubmed_upper.replace('-', ' ').split() if len(w) > 2]
            jcr_words = [w for w in jcr_upper.replace('-', ' ').split() if len(w) > 2]
            
            if not pubmed_words or not jcr_words:
                return 0
            
            # 策略1: 计算单词匹配度
            matched_words = 0
            for pw in pubmed_words:
                for jw in jcr_words:
                    # 完全匹配
                    if pw == jw:
                        matched_words += 1
                        break
                    # 前缀匹配（至少3个字符）
                    elif len(pw) >= 3 and len(jw) >= 3:
                        min_len = min(len(pw), len(jw))
                        prefix_len = 0
                        for i in range(min_len):
                            if pw[i] == jw[i]:
                                prefix_len += 1
                            else:
                                break
                        if prefix_len >= 3:  # 前3个字符相同
                            matched_words += 0.7
                            break
            
            word_score = matched_words / max(len(pubmed_words), len(jcr_words))
            
            # 策略2: 整体字符串相似度（使用 difflib）
            string_score = difflib.SequenceMatcher(None, pubmed_upper, jcr_upper).ratio()
            
            # 综合得分（单词匹配权重更高）
            final_score = word_score * 0.7 + string_score * 0.3
            
            return final_score
        
        # 开始补充匹配
        print("\n开始智能模糊匹配...")
        match_results = {
            "jcr_matched": 0,
            "csa_matched": 0,
            "both_matched": 0,
            "still_unmatched": 0
        }
        
        matched_details = []
        
        for row_idx, pubmed_journal in unmatched_rows:
            best_jcr_match = None
            best_jcr_score = 0
            best_csa_match = None
            best_csa_score = 0
            
            # 在 JCR 数据库中查找最佳匹配
            for jcr_key, jcr_data in jcr_journals.items():
                # 对全称和缩略名都进行匹配
                score_full = fuzzy_match_score(pubmed_journal, jcr_data["full"])
                score_abbr = fuzzy_match_score(pubmed_journal, jcr_data["abbr"]) if jcr_data["abbr"] else 0
                score = max(score_full, score_abbr)
                
                if score > best_jcr_score and score >= min_similarity:
                    best_jcr_score = score
                    best_jcr_match = jcr_data
            
            # 在 CSA 数据库中查找最佳匹配
            for csa_key, csa_data in csa_journals.items():
                score_full = fuzzy_match_score(pubmed_journal, csa_data["full"])
                score_abbr = fuzzy_match_score(pubmed_journal, csa_data["abbr"]) if csa_data["abbr"] else 0
                score = max(score_full, score_abbr)
                
                if score > best_csa_score and score >= min_similarity:
                    best_csa_score = score
                    best_csa_match = csa_data
            
            # 更新 Excel
            jcr_matched = False
            csa_matched = False
            
            if best_jcr_match:
                ws.cell(row=row_idx, column=self.excel_property_dic["IF"]).value = best_jcr_match["IF"]
                ws.cell(row=row_idx, column=self.excel_property_dic["Quartile"]).value = best_jcr_match["Quartile"]
                jcr_matched = True
                match_results["jcr_matched"] += 1
            
            if best_csa_match:
                ws.cell(row=row_idx, column=self.excel_property_dic["JCR_Quartile"]).value = best_csa_match["CSA_Quartile"]
                ws.cell(row=row_idx, column=self.excel_property_dic["Top"]).value = best_csa_match["Top"]
                ws.cell(row=row_idx, column=self.excel_property_dic["OA"]).value = best_csa_match["OA"]
                csa_matched = True
                match_results["csa_matched"] += 1
            
            if jcr_matched and csa_matched:
                match_results["both_matched"] += 1
            
            if not jcr_matched and not csa_matched:
                match_results["still_unmatched"] += 1
            
            # 记录匹配详情
            if jcr_matched or csa_matched:
                detail = f"  {pubmed_journal[:40]}"
                if jcr_matched:
                    detail += f" -> JCR: {best_jcr_match['abbr'] or best_jcr_match['full'][:30]} (相似度:{best_jcr_score:.2f})"
                if csa_matched:
                    detail += f" -> CSA: {best_csa_match['abbr'] or best_csa_match['full'][:30]} (相似度:{best_csa_score:.2f})"
                matched_details.append(detail)
        
        # 保存文件
        wb.save(excel_path)
        
        # 打印结果
        print("\n" + "="*70)
        print("补充匹配完成")
        print("="*70)
        print(f"\n原未匹配记录数: {len(unmatched_rows)}")
        print(f"\n补充匹配结果:")
        print(f"  - JCR 补充匹配成功: {match_results['jcr_matched']} ({match_results['jcr_matched']/len(unmatched_rows)*100:.1f}%)")
        print(f"  - CSA 补充匹配成功: {match_results['csa_matched']} ({match_results['csa_matched']/len(unmatched_rows)*100:.1f}%)")
        print(f"  - 完全匹配（JCR+CSA）: {match_results['both_matched']} ({match_results['both_matched']/len(unmatched_rows)*100:.1f}%)")
        print(f"  - 仍未匹配: {match_results['still_unmatched']} ({match_results['still_unmatched']/len(unmatched_rows)*100:.1f}%)")
        
        if matched_details:
            print(f"\n匹配详情（前20个）:")
            for detail in matched_details[:20]:
                print(detail)
        
        print("\n" + "="*70)
        print(f"已更新文件: {excel_path}")
        print("="*70)

def download_pdf(self, excel_path, pdf_savepath, IF_cutoff):
        '''
        try to download paper which IF higher than cutoff
        warning: very low successful rate
        '''
        
        wb = openpyxl.load_workbook(excel_path)
        ws = wb["Sheet"]
        base_url = "https://sci-hub.tw/"
        success_count = 0
        for cur_row in trange(2, ws.max_row+1, desc="downloading pdf"):
            IF = ws.cell(row=cur_row, column=self.excel_property_dic["IF"]).value
            pmid = ws.cell(row=cur_row, column=self.excel_property_dic["PMID"]).value
            title = ws.cell(row=cur_row, column=self.excel_property_dic["TI"]).value
            if IF=="Unknow" or float(IF)<IF_cutoff:
                continue

            file_name = pdf_savepath+pmid+"_"+title+".pdf"
            try:
                doi = ws.cell(row=cur_row, column=self.excel_property_dic["LID"]).value.split(" ")[0]
                url = base_url + doi
                getpage = requests.get(url, verify=True)
                getpage_soup = BeautifulSoup(getpage.text, "html.parser")
                src = getpage_soup.find("iframe", src=True).get_attribute_list("src")[0]
                response = requests.get("https:"+src, verify=True)
                f = open(file_name, "wb+")
                f.write(response.content)
                f.close()
                success_count += 1
            except:
                pass
            time.sleep(1)
        print("successful download: {}".format(success_count))
