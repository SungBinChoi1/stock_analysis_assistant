"""
네이버 상승종목 대시보드
Streamlit 기반 인터랙티브 대시보드

실행 방법:
streamlit run stock_dashboard.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import glob
import os
import time


# 페이지 설정
st.set_page_config(
    page_title="상승종목 대시보드",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 스타일
st.markdown("""
<style>
    .main-header {
        font-size: 48px;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 30px;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
    }
    .st-emotion-cache-16idsys p {
        font-size: 18px;
    }
</style>
""", unsafe_allow_html=True)


@st.cache_data
def load_latest_data():
    """최신 데이터 파일 로드"""
    data_files = glob.glob('data/rising_stocks_*.xlsx')

    if not data_files:
        return None, None

    latest_file = max(data_files, key=os.path.getctime)
    df = pd.read_excel(latest_file, sheet_name='상승종목')

    # 날짜 추출
    file_date = os.path.basename(latest_file).replace('rising_stocks_', '').replace('.xlsx', '')

    return df, file_date


def main():
    # 헤더
    st.markdown('<div class="main-header">📈 상승종목 대시보드</div>', unsafe_allow_html=True)

    # 데이터 로드
    df, file_date = load_latest_data()

    if df is None:
        st.error("❌ 데이터 파일이 없습니다. 먼저 데이터를 수집하세요.")
        return

    st.success(f"✅ 데이터 로드 완료: {file_date}")

    # 사이드바
    with st.sidebar:
        st.header("🎛️ 필터")

        # 자동 새로고침 옵션
        st.write("---")
        auto_refresh = st.checkbox("🔄 자동 새로고침", value=False)
        if auto_refresh:
            refresh_interval = st.slider("새로고침 간격 (초)", 10, 300, 60)
            st.caption(f"⏱️ {refresh_interval}초마다 데이터 새로고침")
            time.sleep(refresh_interval)
            st.rerun()

        st.write("---")

        # 시장 구분
        markets = st.multiselect(
            "시장",
            options=df['시장구분'].unique() if '시장구분' in df.columns else [],
            default=df['시장구분'].unique() if '시장구분' in df.columns else []
        )

        # 상태
        statuses = st.multiselect(
            "상태",
            options=df['상태'].unique() if '상태' in df.columns else [],
            default=['정상'] if '상태' in df.columns and '정상' in df['상태'].unique() else []
        )

        # 등락률 범위
        if '등락률' in df.columns:
            rate_range = st.slider(
                "등락률 범위 (%)",
                float(df['등락률'].min()),
                float(df['등락률'].max()),
                (float(df['등락률'].min()), float(df['등락률'].max()))
            )

        # ROE 필터
        if 'ROE' in df.columns:
            roe_min = st.number_input("최소 ROE (%)", value=0.0, step=1.0)

        # PBR 필터
        if 'PBR' in df.columns:
            pbr_max = st.number_input("최대 PBR", value=100.0, step=0.5)

    # 필터 적용
    filtered_df = df.copy()

    if markets and '시장구분' in df.columns:
        filtered_df = filtered_df[filtered_df['시장구분'].isin(markets)]

    if statuses and '상태' in df.columns:
        filtered_df = filtered_df[filtered_df['상태'].isin(statuses)]

    if '등락률' in df.columns:
        filtered_df = filtered_df[
            (filtered_df['등락률'] >= rate_range[0]) &
            (filtered_df['등락률'] <= rate_range[1])
        ]

    if 'ROE' in df.columns:
        filtered_df = filtered_df[filtered_df['ROE'] >= roe_min]

    if 'PBR' in df.columns:
        filtered_df = filtered_df[filtered_df['PBR'] <= pbr_max]

    # 주요 지표
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("📊 총 종목", f"{len(filtered_df):,}개")

    with col2:
        if '등락률' in filtered_df.columns:
            avg_rate = filtered_df['등락률'].mean()
            st.metric("📈 평균 등락률", f"{avg_rate:.2f}%")

    with col3:
        if 'ROE' in filtered_df.columns:
            avg_roe = filtered_df['ROE'].dropna().mean()
            st.metric("💰 평균 ROE", f"{avg_roe:.2f}%")

    with col4:
        if '등락률' in filtered_df.columns:
            limit_up = len(filtered_df[filtered_df['등락률'] >= 29.5])
            st.metric("🔥 상한가", f"{limit_up}개")

    # 탭
    tab1, tab2, tab3, tab4 = st.tabs(["📊 차트", "📋 테이블", "🔍 상세검색", "📰 뉴스"])

    with tab1:
        st.subheader("📊 시각화")

        col1, col2 = st.columns(2)

        with col1:
            # 등락률 분포
            if '등락률' in filtered_df.columns and '등락구분' in filtered_df.columns:
                fig = px.histogram(
                    filtered_df,
                    x='등락률',
                    color='등락구분',
                    title='등락률 분포',
                    nbins=30,
                    color_discrete_map={
                        '상한가': '#ff0000',
                        '상승': '#ff6b6b',
                        '보합': '#95a5a6',
                        '하락': '#3498db',
                        '하한가': '#2980b9'
                    }
                )
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)

        with col2:
            # 시장별 분포
            if '시장구분' in filtered_df.columns:
                market_counts = filtered_df['시장구분'].value_counts()
                fig = px.pie(
                    values=market_counts.values,
                    names=market_counts.index,
                    title='시장별 분포',
                    color_discrete_map={'코스피': '#3498db', '코스닥': '#e74c3c'}
                )
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)

        # ROE vs PBR 산점도
        if 'ROE' in filtered_df.columns and 'PBR' in filtered_df.columns:
            st.subheader("ROE vs PBR")

            scatter_df = filtered_df[
                (filtered_df['ROE'].notna()) &
                (filtered_df['PBR'].notna()) &
                (filtered_df['ROE'] > 0) &
                (filtered_df['PBR'] > 0) &
                (filtered_df['ROE'] < 100) &
                (filtered_df['PBR'] < 10)
            ].copy()

            if len(scatter_df) > 0:
                fig = px.scatter(
                    scatter_df,
                    x='PBR',
                    y='ROE',
                    size='시가총액' if '시가총액' in scatter_df.columns else None,
                    color='등락률' if '등락률' in scatter_df.columns else None,
                    hover_data=['종목명', '현재가'] if '종목명' in scatter_df.columns else None,
                    title='ROE vs PBR (가치주 찾기)',
                    color_continuous_scale='RdYlGn'
                )
                fig.add_hline(y=15, line_dash="dash", line_color="green", annotation_text="ROE 15%")
                fig.add_vline(x=1.5, line_dash="dash", line_color="blue", annotation_text="PBR 1.5")
                fig.update_layout(height=500)
                st.plotly_chart(fig, use_container_width=True)

        # 업종별 평균
        if '업종' in filtered_df.columns and '등락률' in filtered_df.columns:
            st.subheader("업종별 평균 등락률")

            sector_df = filtered_df.groupby('업종')['등락률'].mean().sort_values(ascending=False).head(15)

            fig = px.bar(
                x=sector_df.values,
                y=sector_df.index,
                orientation='h',
                title='업종별 평균 등락률 TOP 15',
                labels={'x': '평균 등락률 (%)', 'y': '업종'}
            )
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)

    with tab2:
        st.subheader("📋 데이터 테이블")

        # 컬럼 선택
        display_cols = st.multiselect(
            "표시할 컬럼 선택",
            options=filtered_df.columns.tolist(),
            default=['종목명', '종목코드', '현재가', '등락률', '시장구분', '업종', 'ROE', 'PBR']
                    if all(col in filtered_df.columns for col in ['종목명', '종목코드', '현재가', '등락률'])
                    else filtered_df.columns[:8].tolist()
        )

        if display_cols:
            # 정렬
            sort_col = st.selectbox("정렬 기준", display_cols, index=3 if '등락률' in display_cols else 0)
            sort_order = st.radio("정렬 순서", ["내림차순", "오름차순"], horizontal=True)

            display_df = filtered_df[display_cols].copy()
            display_df = display_df.sort_values(
                by=sort_col,
                ascending=(sort_order == "오름차순")
            )

            # 테이블 표시
            st.dataframe(
                display_df,
                use_container_width=True,
                height=600
            )

            # 다운로드
            csv = display_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 CSV 다운로드",
                data=csv,
                file_name=f'filtered_stocks_{datetime.now().strftime("%Y%m%d")}.csv',
                mime='text/csv'
            )

    with tab3:
        st.subheader("🔍 상세 검색")

        col1, col2 = st.columns(2)

        with col1:
            st.write("**재무 조건**")

            roe_range = st.slider("ROE 범위", 0, 100, (0, 100)) if 'ROE' in filtered_df.columns else None
            pbr_range = st.slider("PBR 범위", 0.0, 10.0, (0.0, 10.0), step=0.1) if 'PBR' in filtered_df.columns else None
            per_range = st.slider("PER 범위", 0, 100, (0, 100)) if 'PER' in filtered_df.columns else None

        with col2:
            st.write("**시장 조건**")

            cap_range = st.slider(
                "시가총액 (억원)",
                0,
                int(filtered_df['시가총액'].max() / 100000000) if '시가총액' in filtered_df.columns else 10000,
                (0, int(filtered_df['시가총액'].max() / 100000000) if '시가총액' in filtered_df.columns else 10000)
            ) if '시가총액' in filtered_df.columns else None

        # 검색 버튼
        if st.button("🔍 검색", type="primary"):
            search_df = filtered_df.copy()

            if roe_range and 'ROE' in search_df.columns:
                search_df = search_df[(search_df['ROE'] >= roe_range[0]) & (search_df['ROE'] <= roe_range[1])]

            if pbr_range and 'PBR' in search_df.columns:
                search_df = search_df[(search_df['PBR'] >= pbr_range[0]) & (search_df['PBR'] <= pbr_range[1])]

            if per_range and 'PER' in search_df.columns:
                search_df = search_df[(search_df['PER'] >= per_range[0]) & (search_df['PER'] <= per_range[1])]

            if cap_range and '시가총액' in search_df.columns:
                search_df = search_df[
                    (search_df['시가총액'] / 100000000 >= cap_range[0]) &
                    (search_df['시가총액'] / 100000000 <= cap_range[1])
                ]

            st.write(f"**검색 결과: {len(search_df)}개 종목**")

            if len(search_df) > 0:
                st.dataframe(
                    search_df[['종목명', '종목코드', '현재가', '등락률', 'ROE', 'PBR', 'PER', '시가총액']
                              if all(col in search_df.columns for col in ['종목명', '종목코드', '현재가', '등락률', 'ROE', 'PBR'])
                              else search_df.columns[:8]],
                    use_container_width=True,
                    height=400
                )

    with tab4:
        st.subheader("📰 최근 뉴스 & 종목정보")

        # 종목 선택
        if '종목명' in filtered_df.columns:
            selected_stock = st.selectbox(
                "종목 선택",
                filtered_df['종목명'].tolist()
            )

            stock_row = filtered_df[filtered_df['종목명'] == selected_stock].iloc[0]

            # 기본 정보
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("현재가", f"{stock_row['현재가']:,}원" if '현재가' in stock_row else "N/A")
            col2.metric("등락률", f"{stock_row['등락률']:.2f}%" if '등락률' in stock_row else "N/A")
            col3.metric("시장", stock_row['시장구분'] if '시장구분' in stock_row else "N/A")
            col4.metric("업종", stock_row['업종'] if '업종' in stock_row else "N/A")

            # 종목설명 추가
            if '종목설명' in stock_row and stock_row['종목설명']:
                st.write("---")
                st.write("**📝 종목 설명**")
                st.info(stock_row['종목설명'])

            st.write("---")

            # 뉴스
            st.write("**📰 최근 뉴스**")
            news_found = False
            for i in range(1, 4):
                news_col = f'뉴스{i}'
                date_col = f'뉴스{i}_일자'
                link_col = f'뉴스{i}_링크'

                if news_col in stock_row and stock_row[news_col]:
                    news_found = True
                    with st.expander(f"[{stock_row[date_col] if date_col in stock_row else ''}] {stock_row[news_col]}"):
                        if link_col in stock_row and stock_row[link_col]:
                            st.markdown(f"[🔗 기사 링크]({stock_row[link_col]})")

            if not news_found:
                st.caption("최근 뉴스가 없습니다.")

            # 공시
            st.write("**📋 최근 공시**")
            notice_found = False
            for i in range(1, 4):
                notice_col = f'공시{i}'
                date_col = f'공시{i}_일자'
                link_col = f'공시{i}_링크'

                if notice_col in stock_row and stock_row[notice_col]:
                    notice_found = True
                    with st.expander(f"[{stock_row[date_col] if date_col in stock_row else ''}] {stock_row[notice_col]}"):
                        if link_col in stock_row and stock_row[link_col]:
                            st.markdown(f"[🔗 공시 링크]({stock_row[link_col]})")

            if not notice_found:
                st.caption("최근 공시가 없습니다.")

            # IR
            st.write("**🏢 최근 IR**")
            ir_found = False
            for i in range(1, 4):
                ir_col = f'IR{i}'
                date_col = f'IR{i}_일자'

                if ir_col in stock_row and stock_row[ir_col]:
                    ir_found = True
                    st.write(f"- [{stock_row[date_col] if date_col in stock_row else ''}] {stock_row[ir_col]}")

            if not ir_found:
                st.caption("최근 IR이 없습니다.")

    # 푸터
    st.markdown("---")
    st.caption(f"데이터 수집: {file_date} | 총 {len(df)}개 종목")


if __name__ == "__main__":
    main()