import ddddocr

ocr = ddddocr.DdddOcr()
with open('code.png', 'rb') as f:
    img_bytes = f.read()
res = ocr.classification(img_bytes)
print('识别出的验证码为：' + res)
