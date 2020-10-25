處理檔案命名為 source.xls 輸出檔案為 out.csv

執行指令為 npm start

檔案內容設定說明
SHEETNAME 為表名(現在為「資料」)
RANGE 為資料處理範圍 `const RANGE = 'A2:M992'` A2, M992 為範圍內的左上&右下角
RANGE 設愈大所需吃的記憶體愈大，故多設一點可以，不然會程式讀 excel 時會吃幾萬筆的資料，
故程式內先瑣定某個範圍，如需擴展請記的調整該值。
現況排除資料的檢查為 `checkUnit` 此 function 如需增設條件可至該 function 內增補
預設第一欄呈現為「尺碼」(變數為 `const SIZE=‘尺碼’`

資料欄位(source.xlsx)
A:通用商品編號(Common product number)
B:商品編號
C:尺寸名稱(Size name)
D:
E:複製過來
F:手動整理的結果
G 之後為規格表
A~E 順序請勿更動
所有的丈量尺寸值以「G 欄」之後的值為基準

FAQ
如需增補丈量尺寸，該怎做?
Ans. 請修正 `RANGE='範圍'` 參數即可 eg. A2M10 -> A2N10 A2, N10 為範圍內的左上&右下角

資料筆數短缺，該怎做?
Ans. 請修正 `RANGE='範圍'` 參數即可 eg. A2M10 -> A2M100 A2, M100 為範圍內的左上&右下角
