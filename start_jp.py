import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import re
from datetime import datetime, timedelta
import io
from typing import List, Dict, Any

# Page configuration
st.set_page_config(
    page_title="スケジュール変換ツール",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        margin-bottom: 2rem;
    }
    .step-container {
        border: 2px solid #667eea;
        border-radius: 10px;
        padding: 1rem;
        margin: 1rem 0;
        background-color: #f8f9fa;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

class ScheduleConverter:
    def __init__(self):
        self.parsed_schedule = []
        
    def parse_schedule_text(self, text: str) -> List[Dict[str, Any]]:
        """Parse Japanese schedule text format"""
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        schedule = []
        
        # Extract year and month from date range
        year_month_match = re.search(r'(\d{4})年(\d{1,2})月', lines[0])
        year = int(year_month_match.group(1)) if year_month_match else datetime.now().year
        month = int(year_month_match.group(2)) if year_month_match else datetime.now().month
        
        current_date = None
        current_full_date = None
        i = 0
        
        # Find where the actual schedule data starts
        while i < len(lines):
            line = lines[i]
            date_match = re.match(r'^(\d{1,2})\([日月火水木金土]\)$', line)
            if date_match:
                break
            i += 1
        
        while i < len(lines):
            line = lines[i]
            
            # Check for date line (format: "23(月)")
            date_match = re.match(r'^(\d{1,2})\([日月火水木金土]\)$', line)
            if date_match:
                current_date = line
                day_num = int(date_match.group(1))
                current_full_date = datetime(year, month, day_num)
                i += 1
                continue
            
            # Check for content in English parentheses (no time/location)
            if current_full_date and line.startswith('(') and line.endswith(')'):
                activity_text = line[1:-1].strip()
                schedule.append({
                    'date': current_date,
                    'full_date': current_full_date,
                    'time': '',
                    'location': '',
                    'activity': activity_text,
                    'has_all_data': False
                })
                i += 1
                continue
            
            # Check for time and activity line (format: "08:50 川口本部")
            time_activity_match = re.match(r'^(\d{1,2}):(\d{2})\s+(.+)$', line)
            if time_activity_match and current_full_date:
                time = f"{time_activity_match.group(1)}:{time_activity_match.group(2)}"
                activity_text = time_activity_match.group(3).strip()
                
                # Parse location and activity
                location = ''
                activity = ''
                
                if '(' in activity_text and ')' in activity_text:
                    # Text in parentheses is activity
                    activity = re.sub(r'[()]', '', activity_text)
                    location = ''
                elif activity_text == '社用車帰宅':
                    location = '社用車帰宅'
                    activity = ''
                else:
                    # Split by space
                    parts = re.split(r'[　\s]+', activity_text)
                    location = parts[0] if parts else ''
                    activity = ' '.join(parts[1:]) if len(parts) > 1 else ''
                
                # Handle special case
                if not activity and location:
                    if any(keyword in location for keyword in ['打合せ', '会議', '見学', '参加', '食事', '手配', '対応']):
                        activity = location
                        location = ''
                
                # Determine if this entry has all data
                has_all_data = bool(time and (location or activity))
                
                schedule.append({
                    'date': current_date,
                    'full_date': current_full_date,
                    'time': time,
                    'location': location,
                    'activity': activity,
                    'has_all_data': has_all_data
                })
            
            i += 1
        
        return schedule
    
    def generate_filename(self, text: str) -> str:
        """Generate filename from schedule date range"""
        date_range_match = re.search(
            r'(\d{4})年(\d{1,2})月(\d{1,2})日\([日月火水木金土]\)\s*～\s*(\d{4})年(\d{1,2})月(\d{1,2})日\([日月火水木金土]\)',
            text
        )
        
        if date_range_match:
            start_year = date_range_match.group(1)
            start_month = date_range_match.group(2).zfill(2)
            start_day = date_range_match.group(3).zfill(2)
            end_year = date_range_match.group(4)
            end_month = date_range_match.group(5).zfill(2)
            end_day = date_range_match.group(6).zfill(2)
            
            return f"{start_year}{start_month}{start_day}to{end_year}{end_month}{end_day}.xlsx"
        
        return f"weekly_schedule_{datetime.now().strftime('%Y%m%d')}.xlsx"
    
    def create_excel_file(self, template_file, schedule_data: List[Dict]) -> bytes:
        """Create Excel file with schedule data"""
        wb = load_workbook(template_file)
        ws = wb.active
        
        if not schedule_data:
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            return excel_buffer.getvalue()
        
        # Extract start date from first entry
        first_entry = min(schedule_data, key=lambda x: x['full_date'])
        start_date = first_entry['full_date']
        
        # Generate all 7 days of the week
        all_dates = []
        for i in range(7):
            date = start_date + timedelta(days=i)
            all_dates.append(date)
        
        # Group schedule by date
        schedule_by_date = {}
        for item in schedule_data:
            date_key = item['full_date'].strftime('%Y-%m-%d')
            if date_key not in schedule_by_date:
                schedule_by_date[date_key] = []
            schedule_by_date[date_key].append(item)
        
        # Sort entries: complete entries first
        for date_key in schedule_by_date:
            schedule_by_date[date_key].sort(key=lambda x: (not x.get('has_all_data', True), x['time'] or '99:99'))
        
        # Japanese day names
        day_names = ['月', '火', '水', '木', '金', '土', '日']
        
        # Populate Excel sheet
        for day_idx, date in enumerate(all_dates):
            date_key = date.strftime('%Y-%m-%d')
            day_entries = schedule_by_date.get(date_key, [])
            
            # Calculate starting row (6-row blocks)
            day_start_row = 7 + (day_idx * 6)
            
            # Get day of week
            weekday_idx = date.weekday()
            day_name = day_names[weekday_idx]
            
            # Format the date column
            for row_offset in range(6):
                current_row = day_start_row + row_offset
                
                if row_offset == 0:
                    # First row: Day name in parentheses
                    ws.cell(row=current_row, column=1, value=f"({day_name})")
                elif row_offset == 1:
                    # Second row: Full date
                    ws.cell(row=current_row, column=1, value=date.strftime('%Y/%m/%d'))
                else:
                    # Rows 3-6: Leave blank
                    ws.cell(row=current_row, column=1, value="")
            
            # Process entries for this day
            row_index = day_start_row
            for entry in day_entries:
                # Set time
                if entry['time']:
                    time_parts = entry['time'].split(':')
                    time_value = datetime(1899, 12, 30, int(time_parts[0]), int(time_parts[1]))
                    ws.cell(row=row_index, column=2, value=time_value)
                else:
                    ws.cell(row=row_index, column=2, value="")
                
                # Set location (remove parentheses)
                location_clean = re.sub(r'[()]', '', entry['location']) if entry['location'] else ''
                ws.cell(row=row_index, column=3, value=location_clean)
                
                # Set activity (remove parentheses)
                activity_clean = re.sub(r'[()]', '', entry['activity']) if entry['activity'] else ''
                ws.cell(row=row_index, column=4, value=activity_clean)
                
                row_index += 1
                
                # Don't exceed the 6-row block
                if row_index >= day_start_row + 6:
                    break
        
        # Save to bytes
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        return excel_buffer.getvalue()

def main():
    st.markdown('<h1 class="main-header">📅 スケジュール変換ツール</h1>', unsafe_allow_html=True)
    
    # Initialize converter
    converter = ScheduleConverter()
    
    # Sidebar
    with st.sidebar:
        st.markdown("## 📋 使用方法")
        st.markdown("""
        1. **Excelテンプレートをアップロード**
        2. **スケジュールテキストを貼り付け**
        3. **解析とプレビュー**
        4. **Excelファイルをダウンロード**
        """)
        
        st.markdown("## 🔧 機能")
        st.markdown("""
        - ✅ 日本語テキストの解析
        - ✅ Excelフォーマットの保持
        - ✅ 6行ブロック構造
        - ✅ 英語括弧対応 `(内容)`
        - ✅ 完全データの優先表示
        """)
    
    # Main content
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        st.markdown("### 📄 ステップ1: Excelテンプレート")
        
        excel_template = st.file_uploader(
            "Excelテンプレートファイル (.xlsx) を選択",
            type=['xlsx']
        )
        
        if excel_template:
            st.markdown('<div class="success-box">✅ テンプレート読み込み完了</div>', unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        st.markdown("### 📱 ステップ2: スケジュールデータ")
        
        sample_text = """ここにテキストを貼り付けてください。"""
        
        schedule_text = st.text_area(
            "スケジュールテキストを貼り付け:",
            value=sample_text,
            height=400
        )
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="step-container">', unsafe_allow_html=True)
        st.markdown("### 🔍 ステップ3: 解析とプレビュー")
        
        if st.button("🔍 スケジュールを解析", type="primary"):
            if schedule_text.strip():
                try:
                    converter.parsed_schedule = converter.parse_schedule_text(schedule_text)
                    st.session_state.parsed_schedule = converter.parsed_schedule
                    st.session_state.schedule_text = schedule_text
                    
                    complete_entries = sum(1 for item in converter.parsed_schedule if item.get('has_all_data', True))
                    incomplete_entries = len(converter.parsed_schedule) - complete_entries
                    
                    st.markdown(f'<div class="success-box">✅ {len(converter.parsed_schedule)}件を解析完了！<br/>完全データ: {complete_entries}件 | 部分データ: {incomplete_entries}件</div>', unsafe_allow_html=True)
                except Exception as e:
                    st.markdown(f'<div class="error-box">❌ エラー: {str(e)}</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="error-box">❌ テキストを入力してください</div>', unsafe_allow_html=True)
        
        # Display preview
        if hasattr(st.session_state, 'parsed_schedule') and st.session_state.parsed_schedule:
            st.markdown("#### 📋 プレビュー:")
            
            preview_data = []
            for item in st.session_state.parsed_schedule:
                priority = "🔴 完全" if item.get('has_all_data', True) else "🟡 部分"
                location_clean = re.sub(r'[()]', '', item['location']) if item['location'] else '-'
                activity_clean = re.sub(r'[()]', '', item['activity']) if item['activity'] else '-'
                
                preview_data.append({
                    '月日': item['full_date'].strftime('%m/%d (%a)'),
                    '優先度': priority,
                    'AM/PM': item['time'] or '-',
                    '訪問先': location_clean,
                    '面談内容': activity_clean
                })
            
            df = pd.DataFrame(preview_data)
            st.dataframe(df, use_container_width=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Generate Excel
    st.markdown('<div class="step-container">', unsafe_allow_html=True)
    st.markdown("### 📊 ステップ4: Excelファイル生成")
    
    col3, col4 = st.columns([1, 1])
    
    with col3:
        if st.button("📊 Excelファイルを生成", type="primary", use_container_width=True):
            if not excel_template:
                st.error("Excelテンプレートをアップロードしてください")
            elif not hasattr(st.session_state, 'parsed_schedule') or not st.session_state.parsed_schedule:
                st.error("先にスケジュールを解析してください")
            else:
                try:
                    excel_data = converter.create_excel_file(excel_template, st.session_state.parsed_schedule)
                    filename = converter.generate_filename(st.session_state.schedule_text)
                    
                    st.download_button(
                        label="💾 Excelファイルをダウンロード",
                        data=excel_data,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    st.success(f"✅ '{filename}' 生成完了！")
                    
                except Exception as e:
                    st.error(f"❌ エラー: {str(e)}")
    
    with col4:
        st.markdown('<div class="info-box">', unsafe_allow_html=True)
        st.markdown("**💡 対応形式:**")
        st.markdown("""
        - **完全予定**: `08:50 川口本部`
        - **活動のみ**: `(梱包資材購入)`
        - **括弧内容**: `15:00 (VE会議)`
        """)
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()