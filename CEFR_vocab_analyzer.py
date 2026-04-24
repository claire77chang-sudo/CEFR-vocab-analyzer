import os
import re
import pandas as pd
from collections import Counter #計算出現頻率

#文字處理
import nltk
from nltk.stem import WordNetLemmatizer #還原詞性
from nltk.corpus import stopwords #載入停用詞

#難度分析
from cefrpy import CEFRAnalyzer

#檔案載入、清理
from ebooklib import epub, ITEM_DOCUMENT
from bs4 import BeautifulSoup #拆解網頁格式

#輸出成Excel格式
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment #Excel儲存格

#API
import requests #HTTP請求
import time #控制API請求時間

#程式錯誤預防(用於中斷程式)
import sys



#下載單字清洗比對資料庫
nltk.download('wordnet', quiet=True)
nltk.download('omw-1.4', quiet=True)
nltk.download('stopwords', quiet=True)

lemmatizer = WordNetLemmatizer()
cefr_analyzer = CEFRAnalyzer()
#把停用詞也配合還原成原形動詞，方便之後過濾
raw_stop_words = set(stopwords.words('english'))
stop_words = {lemmatizer.lemmatize(w, pos='v') for w in raw_stop_words}
stop_words.update({lemmatizer.lemmatize(w, pos='n') for w in stop_words})


def check_file_lock(file_path):
    if os.path.exists(file_path):
        try:
            # 嘗試重新命名自己，檢查檔案是否被開啟
            os.rename(file_path, file_path)
        except OSError:
            print(f"❌ 錯誤：檔案 [{os.path.basename(file_path)}] 正在開啟中！請先關閉 Excel 再繼續。")
            sys.exit() # 直接中斷，不跑後面的同步


class DictionaryLookup:
    #處理線上 API 查詢並維護執行期間快取
    def __init__(self, local_csv_path=None):
        self.api_url = "https://api.dictionaryapi.dev/api/v2/entries/en/"
        self.cache = {}
        self.local_dict = {}
        
        # 載入本地 ECDICT 
        if local_csv_path and os.path.exists(local_csv_path):
            try:
                print(f"📚 載入離線辭庫: {local_csv_path}")
                df = pd.read_csv(local_csv_path, usecols=['word', 'phonetic', 'definition'], encoding='utf-8')
                df['word'] = df['word'].astype(str).str.lower()
                df = df.drop_duplicates(subset=['word'])
                self.local_dict = df.set_index('word').to_dict('index')
            except Exception as e:
                print(f"⚠️ 離線辭庫載入失敗: {e}")

    def get_info(self, word):
        word = str(word).lower().strip()

        if word in self.cache:
            return self.cache[word] # 當字串[word]再快取內，直接回傳快取中的字串
        
        # 第一層：查本地 ECDICT
        if word in self.local_dict:
            item = self.local_dict[word]
            res = {
                'Phonetic': str(item.get('phonetic', '')),
                'Definition': str(item.get('definition', '')).replace('\\n', ' '),
                'Source': 'Local'
            }
            self.cache[word] = res
            return res

        # 第二層：查線上 API
        try:
            time.sleep(0.2)
            response = requests.get(self.api_url + word, timeout=5)
            if response.status_code == 200:
                data = response.json()[0]
            
                # 提取音標
                phonetic = data.get('phonetic', '') or (data.get('phonetics')[0].get('text', '') if data.get('phonetics') else '')      
                # 提取第一個定義
                meanings = data.get('meanings', [{}])[0]
                defs = meanings.get('definitions', [{}])[0]
                
                result = {
                    'Phonetic': phonetic,
                    'Definition': 
                    f"[{meanings.get('partOfSpeech', '')}] {defs.get('definition', '')}",
                    'Source': 'API' 
                }
                self.cache[word] = result 
                return result
        except:
            pass
        
        # 查詢失敗的回傳格式
        fail_res = {'Phonetic': '', 'Definition': 'Not Found', 'Example': '', 'Source': 'None'}
        self.cache[word] = fail_res
        return fail_res

# ==========================================
# 反向同步與 Excel 處理
# ==========================================

def sync_known_words(base_dir, known_xlsx_path):
    #從報告回收標記為 Yes 的單字
    print("\n" + "="*30)
    print("🔍 [DEBUG] 啟動反向同步檢查...")
    
    # 1. 確保總表存在
    if not os.path.exists(known_xlsx_path):
        pd.DataFrame(columns=['Word', 'Note']).to_excel(known_xlsx_path, index=False)
        print(f"📝 建立新總表: {known_xlsx_path}")

    new_learned_words = []

    # 2. 獲取所有報告檔案 (排除 known_words 本身)
    all_files = os.listdir(base_dir)
    report_files = [f for f in all_files if f.endswith("_vocabulary_report.xlsx")]
    
    print(f"📄 偵測到 {len(report_files)} 份報告檔案待掃描")

    for report in report_files:
        report_path = os.path.join(base_dir, report)
        try:
            # 讀取 Excel 內所有分頁
            with pd.ExcelFile(report_path) as xls:
                for sheet_name in xls.sheet_names:
                    if "Summary" in sheet_name: continue
                    df = pd.read_excel(report_path, sheet_name=sheet_name)
                    if 'Is_Known' in df.columns and 'Word' in df.columns:

                        # 確保有 Note 欄位，沒有就補空的
                        if 'Note' not in df.columns: df['Note'] = ""

                        # 嚴格比對: 轉大寫 + 去空格
                        mask = (df['Is_Known'].astype(str).str.strip().str.upper() == 'YES') | (df['Note'].fillna('').astype(str).str.strip() != '')
                        to_move = df[mask][['Word', 'Note']]
                        if not to_move.empty:
                            # 轉換為小寫並存入待合併清單 
                            new_learned_words.append(to_move)

        except Exception as e:
            print(f"  ❌ 讀取 {report} 失敗: {e}")

    # 3. 存入總表
    if new_learned_words:
        try:
            existing_df = pd.read_excel(known_xlsx_path)
            
            # 將所有收集到的小 DataFrame 片段合併
            new_moves = pd.concat(new_learned_words, ignore_index=True)
            
            # 與總表合併
            updated_df = pd.concat([existing_df, new_moves], ignore_index=True)
            
            # 防止重複，並保留最後一次(最新)的筆記
            updated_df.drop_duplicates(subset=['Word'], keep='last', inplace=True)
            
            # 存檔
            updated_df.to_excel(known_xlsx_path, index=False)
            
            # 存檔成功才刪除舊報告
            for report in report_files:
                os.remove(os.path.join(base_dir, report))
                print(f"🗑️ 已完成同步並清理舊檔: {report}")
                
            print(f"✅ 搬家成功！總表已更新。")
        except Exception as e:
            print(f"❌ 更新總表失敗，未刪除舊檔: {e}")

        except Exception as e:
            print(f"❌ 更新總表失敗: {e}")
    else:
        print("🔍 同步完成：沒有在任何報告中找到 'Yes' 標記。")
    print("="*30 + "\n")

def load_known_words_excel(file_path):
    """載入總表，回傳 set"""
    if not os.path.exists(file_path): return set()
    try:
        df = pd.read_excel(file_path)
        return set(df['Word'].dropna().astype(str).str.lower().str.strip())
    except: return set()

# ==========================================
# 【清洗與分析邏輯】
# ==========================================

def clean_and_normalize(raw_content):
    """修正 SyntaxWarning 並執行清洗"""
    text = raw_content.lower()
    text = re.sub(r'[^a-z\s]', ' ', text) 
    raw_words = text.split()
    normalized_words = []
    for w in raw_words:
        lemma = lemmatizer.lemmatize(w, pos='v')
        lemma = lemmatizer.lemmatize(lemma, pos='n')
        if lemma.isalpha() and len(lemma) > 1 and ord(lemma[0]) < 128:
            normalized_words.append(lemma)
    return normalized_words

def calculate_text_difficulty(analysis_results):
    weight_map = {'A1': 1, 'A2': 2, 'B1': 3, 'B2': 4, 'C1': 5, 'C2': 6, 'Unknown': 0}
    total_weighted_score = 0
    total_word_count = 0
    level_counts = Counter()
    for item in analysis_results:
        lvl = item['Level']
        count = item['Count']
        score = weight_map.get(lvl, 0)
        if score > 0:
            total_weighted_score += (score * count)
            total_word_count += count
        level_counts[lvl] += count
    avg_score = total_weighted_score / total_word_count if total_word_count > 0 else 0
    levels = ["A1 (入門)", "A2 (初級)", "B1 (中級)", "B2 (中高)", "C1 (進階)", "C2 (精通)"]
    recommended = levels[min(int(max(0, avg_score - 1)), 5)]

    return {
        'avg_score': round(avg_score, 2),
        'recommended_level': recommended,
        'level_distribution': level_counts
    }

def process_and_enrich(word_list, filter_set):
    # 過濾總表單字與虛詞
    filtered_words = [w for w in word_list if w not in filter_set and w not in stop_words]
    word_counts = Counter(filtered_words)
    unique_level_counts = Counter()
    final_report = []
    print(f"📖 正在分析 {len(word_counts)} 個生字並獲取定義...") # 簡化提示
    
    for word, count in word_counts.items():
        raw_level = cefr_analyzer.get_average_word_level_CEFR(word)
        level_str = str(raw_level) if raw_level is not None else "Unknown"
        unique_level_counts[level_str] += 1
        info = dict_lookup.get_info(word)
        
        final_report.append({
            'Word': word, 
            'Count': count, 
            'Level': level_str, 
            'Phonetic': info['Phonetic'],
            'Definition': info['Definition'],
            'Source': info['Source'],
            'Is_Known': 'No',
            'Note': ''
        })

    print("✅ 生字分析處理完成。")
    sorted_report = sorted(final_report, key=lambda x: x['Count'], reverse=True)
    return sorted_report, unique_level_counts # 多回傳一個統計結果

def export_to_excel(analysis_results, excel_name, summary_data):
    if not analysis_results: return
    df = pd.DataFrame(analysis_results)
    level_order = ['C2', 'C1', 'B2', 'B1', 'A2', 'A1', 'Unknown']
    
    
    try:
        with pd.ExcelWriter(excel_name, engine='openpyxl') as writer:
            # 1. 寫入 Summary 摘要
            summary_content = [
                ["【文本分析摘要】", ""],
                ["總詞彙量", summary_data.get('total_word_count_raw', 0)],
                ["總詞彙量（不含已知）", len(analysis_results)],
                ["加權難度總分 (1-6)", summary_data['avg_score']],
                ["建議閱讀等級", summary_data['recommended_level']],
                ["", ""],
                ["【等級分佈統計】", "出現次數"],
            ]


            # 加入各等級字數
            for lvl in level_order:
                summary_content.append([lvl, summary_data['level_distribution'].get(lvl, 0)])
            
            summary_content.append(["", ""])
            summary_content.append(["【難度分對應標準】", "等級解讀"])
            summary_content.append(["1.0 - 1.4", "A1 極其簡單，幾乎全是基礎詞彙"])
            summary_content.append(["1.5 - 2.4", "A2 基礎入門，開始出現簡單生活詞彙"])
            summary_content.append(["2.5 - 3.4", "B1 進階程度，適合一般英語學習者"])
            summary_content.append(["3.5 - 4.4", "B2 中高難度，具備討論專業話題的門檻"])
            summary_content.append(["4.5 - 5.4", "C1 高級難度，包含大量學術或抽象詞彙"])
            summary_content.append(["5.5 - 6.0", "C2 精通難度，接近母語人士的複雜文學/專業層次"])

            summary_df = pd.DataFrame(summary_content)
            summary_df.to_excel(writer, sheet_name="Summary_報告摘要", index=False, header=False)

            
            pd.DataFrame(summary_content).to_excel(writer, sheet_name="Summary_報告摘要", index=False, header=False)

            # 2. 設定通用格式
            wrap_alignment = Alignment(wrap_text=True, vertical='top')
            dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
            #dv.error = '請選擇 Yes 或 No'; dv.errorTitle = '輸入錯誤'

            # 3. 分頁寫入
            for lvl in level_order:
                # 統一欄位順序：A:Word, B:Count, C:Phonetic, D:Definition, E:Source, F:Is_Known, G:'Note'
                cols = ['Word', 'Count', 'Phonetic', 'Definition',  'Source', 'Is_Known', 'Note']
                lvl_df = df[df['Level'] == lvl][cols]
                
                if not lvl_df.empty:
                    sheet_name = str(lvl)[:31]
                    lvl_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    ws = writer.sheets[sheet_name]
                    
                    # 執行自動換行 (Definition 與 Example 欄位)
                    for row in range(2, len(lvl_df) + 2):
                        ws[f'D{row}'].alignment = wrap_alignment # Definition

                    # 加入下拉選單 (Is_Known 在 F 欄)
                    ws.add_data_validation(dv)
                    dv.add(f'F2:F{len(lvl_df) + 1}')

                    # 欄位寬度調整
                    ws.column_dimensions['A'].width = 18 # Word
                    ws.column_dimensions['D'].width = 50 # Definition (自動換行主體)
                    ws.column_dimensions['E'].width = 10 # Source
                    ws.column_dimensions['F'].width = 12 # Is_Known
                    ws.column_dimensions['G'].width = 30 # Note

            # Summary 頁面美化
            writer.sheets["Summary_報告摘要"].column_dimensions['A'].width = 25
            writer.sheets["Summary_報告摘要"].column_dimensions['B'].width = 50

        print(f"✅ 成功匯出：{os.path.basename(excel_name)}")
    except Exception as e:
        print(f"❌ 匯出失敗 {excel_name}: {e}")


# ==========================================
# 【執行流程控制】
# ==========================================

if __name__ == "__main__":
    # 強制定位工作資料夾為 py 檔所在目錄
    current_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(current_dir)

    # 關鍵：請確認資料夾內檔案名稱確實是 ecdict.csv
    csv_file = os.path.join(current_dir, "ecdict.csv")
    known_path = os.path.join(current_dir, "known_words.xlsx")

    # 1. 檢查總表是否開啟
    check_file_lock(known_path)
    
    # 2. 檢查所有報告檔案是否開啟
    report_files = [f for f in os.listdir(current_dir) if f.endswith("_vocabulary_report.xlsx")]
    for r in report_files:
        check_file_lock(os.path.join(current_dir, r))


    # 增加一個手動檢查，讓你知道到底有沒有讀到檔案
    if not os.path.exists(csv_file):
        print(f"❌ 找不到本地辭庫檔案：{csv_file}程式將切換到慢速 API 模式！請確認本地辭庫ecdict.csv位於程式相同資料夾")


    dict_lookup = DictionaryLookup(local_csv_path=os.path.join(current_dir, "ecdict.csv"))
    known_path = os.path.join(current_dir, "known_words.xlsx")

    # 1. 執行同步
    sync_known_words(current_dir, known_path)

    # 2. 載入已知單字
    MY_KNOWN_WORDS = load_known_words_excel(known_path)
    print(f"✅ 已載入 {len(MY_KNOWN_WORDS)} 個已知單字。\n")

    # 3. 處理檔案
    target_files = [f for f in os.listdir(current_dir) if f.endswith(('.txt', '.epub')) and "_vocabulary_report" not in f]
    if not target_files:
        print(f"📍 在 {current_dir} 中找不到英文檔案。")
    else:
        print(f"🚀 開始處理 {len(target_files)} 個檔案...")
        for file_name in target_files:
            file_path = os.path.join(current_dir, file_name)
            file_ext = os.path.splitext(file_name)[1].lower()
            
            # 讀取
            if file_ext == '.txt':
                try:
                    with open(file_path, 'r', encoding='utf-8') as f: raw_content = f.read()
                except: continue
            else:
                try:
                    book = epub.read_epub(file_path)
                    chapters = []
                    for item in book.get_items_of_type(ITEM_DOCUMENT):
                        chapters.append(BeautifulSoup(item.get_content(), 'html.parser').get_text())
                    raw_content = "\n".join(chapters)
                except: continue

            # 分析與導出
            word_list = clean_and_normalize(raw_content)
            total_raw_count = len(set(word_list))  # 記錄包含已知單字的原始總數
            results, level_unique_stats = process_and_enrich(word_list, MY_KNOWN_WORDS)
            stats  = calculate_text_difficulty(results)
            stats['level_distribution'] = level_unique_stats
            # 將總數存入 summary_stats 字典，方便 export_to_excel 讀取
            stats['total_word_count_raw'] = total_raw_count

            output_name = file_path.replace(file_ext, "_vocabulary_report.xlsx")
            export_to_excel(results, output_name, stats)