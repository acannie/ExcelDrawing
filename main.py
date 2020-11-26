from PIL import Image
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill

# RGBの数値が1桁の場合冒頭に0をつけて2桁にして返す
def addZero(hexRGB):
    if len(hexRGB) == 3:
        return '0x0' + hexRGB[2]
    else:
        return hexRGB

# RGBをカラーコードに変換
def RGBtoColorCode(R, G, B):
    color_code = '{}{}{}'.format(addZero(hex(R)), addZero(
        hex(G)), addZero(hex(B)))  # 0xAA0xBB0xCC の形
    color_code = color_code.replace('0x', '')
    return color_code

# ファイルのインポート
input_file = 'icon.jpg'
im = np.array(Image.open('figure/' + input_file))

# 画像のサイズを取得
img_height = im.shape[0]
img_width = im.shape[1]

# ワークブックを作成
wb = openpyxl.Workbook()
ws = wb.worksheets[0]

n = 3  # 1つのセルの縦幅に対する横幅の大きさ

# ゲージ
print("0                             100%")
print(" ", end="")

now_pacentage = 0

for r in range(1, img_height):
    for c in range(1, int(img_width / n)):
        # 横に並んだnピクセルの平均RGBをセルの背景色とする
        RGB_ave = [0, 0, 0]
        for i in range(n):
            for j in range(im.shape[2]):
                RGB_ave[j] += im[r][n*c + i][j]

        for j in range(im.shape[2]):
            RGB_ave[j] = int(RGB_ave[j]/n)

        color_code = RGBtoColorCode(RGB_ave[0], RGB_ave[1], RGB_ave[2])
        fill = PatternFill(patternType='solid', fgColor=color_code)
        ws.cell(row=r, column=c).fill = fill

    # ゲージを増やす
    if int(r*30 / img_height) > now_pacentage:
        print("■", end="", flush=True)
        now_pacentage = int(r*30 / img_height)

# ファイルのエクスポート
print('')
print("saving...")

output_file = input_file.replace('.jpg', '.xlsx')
wb.save('result/' + output_file)
wb.close()

print("complete!")
