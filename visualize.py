# 금리지표 시각화
# 산금채 금리: 한국산업은행이 발행하는 채권의 금리
# 정기예금 금리: 일정기간 목돈을 한번에 예치하고 만기 시 원금과 이자를 받는 상품 금리
# 정기적금 금리: 매월 일정 금액을 저축하여 만기 시 원금과 함꼐 이자를 받는 상품 금리
# 일반신용대출 금리: 담보 없이 개인 신용도를 기준으로 대출 받는 상품 금리
# 주택담보대출 금리: 주택을 담보로 받는 대출 상품 금리(금리가 낮고, 대출 기간이 장기)

import matplotlib.pyplot as plt
import pandas as pd
from utils.api_data import IMG_DIR
from utils.api_data import OUT2

def indicators_to_png(OUT2):
    with pd.ExcelFile(OUT2) as xlsx:
        for sheet_name in xlsx.sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_name)
            df = df.tail(36)    # 마지막 3개년(36개월) 데이터
            x = df.index
            y = df["DATA_VALUE"]

            y_min = y.min()    # 최솟값
            change = y.iloc[-1] - y.iloc[0]
            color = ("red" if change > 0 else "blue"
                     if change < 0 else "blak")
            fig, ax = plt.subplots(figsize=(9, 3), dpi=100)
            ax.plot(x, y, color=color, linewidth=2)
            # fill_between: 두 지점 사이의 영역을 색으로 채움
            ax.fill_between(x, y, y_min, color=color, alpha=0.10)
            ax.set_axis_off()    # 축 제거
            fig.set_layout_engine("tight")

            fig.savefig(IMG_DIR/f"{sheet_name}.png", bbox_inches='tight', pad_inches=0)


if __name__ == "__main__":
    indicators_to_png(OUT2)