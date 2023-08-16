###---検索くん---###

#-手順-#
#0-1. "python-docx"ライブラリをインストールしよう。
#0-2. "folder_name"にパスを入力しよう。
#1. "find_word"に検索したい語を入力しよう。
#2. "file_name"に探すファイル名を入力しよう。※拡張子なしでおk。
#3. 実行しよう！

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
find_word = "多感覚"
#-----------------------------#
file_name = "内受容感覚"
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#


import docx
folder_name = "C:\\Users\\hikar\\Desktop\\論文メモ"
#例："C:\\Users\\hikar\\Desktop\\論文メモ"
extension_name = ".docx"
docx_name = folder_name + "\\" + file_name + extension_name
doc = docx.Document(docx_name)

check = 0

for i, p in enumerate(doc.paragraphs):
    if i == 0:
        print("|", "段", "|", "位", "|", "文")
        print("|", "落", "|", "置", "|", "章")
    s = p.text.find(find_word)
    if s >= 0:
        print("|", i, "|", s, "|", p.text)
        check = check+1
    
if check == 0:
    print("見つかりませんでした。")


#ページ番号が取得できるともっと便利。