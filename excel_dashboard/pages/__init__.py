"""새 페이지 빌더 모듈 (대회종목·업종·종목분석·뉴스·이벤트).

각 모듈은 build(builder) 하나를 내보냅니다. builder에서 쓸 수 있는 것:
builder.wb, builder.ws[시트명], builder.lo[표이름], builder.table_style, builder.say(),
builder.nav_links(ws, row), builder.HEADER_RIGHT, builder.HEADER_STATUS, 그리고 xl_helpers의 함수.
모듈 등록(호출 순서)은 build_dashboard.py가 맡습니다.
"""
