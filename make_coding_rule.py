'''エクセルファイルからKh coderのコーディングルールファイルを生成する。
数値入力エラ―チェックのバグを修正'''

import tkinter
from tkinter import filedialog
import xlrd
import pandas
import pathlib

#Yes/no判定モジュール
def yes_or_no():
    #全角・半角変換テーブルを作る
    zen = "".join(chr(0xff01 + i) for i in range(94))
    han = "".join(chr(0x21 + i) for i in range(94))

    while True:
        c = input("y/Nを入力してください：").lower()
        #全角入力に対応する
        c.translate(str.maketrans(zen, han))
        #結果を判別する
        if c in ["y", "ye", "yes"]:
            return(True)
        if c in ["n", "no"]:
            return(False)
        else:
            print("入力に誤りがあります。")

#ファイルを出力する
def write_file(list_of_string, path):
    with open(path, "x")as f:
        f.writelines(list_of_string)
    print("■以下の場所にファイルを保存しました")
    print(path)

#保存先とファイル名を確定する
def chose_path(path):
    #同じフォルダに保存するか聞く
    print("■処理が終了しました。先ほどと同じフォルダにファイルを保存しますか？")
    c = yes_or_no()
    if c == True:
        path = pathlib.Path(path)
        folder = path.parent
    else:
    #新しいパスを取得する
        print("ダイアログボックスでファイルを保存するフォルダを選択してください。")
        folder = filedialog.askdirectory()
    file = input("保存するファイル名を入力してください（拡張子は不要です）：")
    path = "{}/{}.txt".format(folder, file)
    return(path)


#書き込むためのデータを作る
def make_data(df_code, df_coding):
    #コード番号のリストを作る
    l_c_n = df_coding.iloc[:, 0].to_list()
    l_c_n = list(map(float, l_c_n))
    #コード名のリストを作る
    l_c_name = df_coding.iloc[:, 1].to_list()
    dict_of_codename = dict(zip(l_c_n, l_c_name))
    #コードする語の辞書を作る
    #df_codeの1列目の名前を変える
    c_name = df_code.columns.values[0]
    df_code = df_code.rename(columns={c_name: "number"})
    #クエリで抽出して辞書を作る
    dict_of_coded_words = {}
    for n in l_c_n:
        dfn = df_code.query("number == @n")
        temp_l = dfn.iloc[:, 1].to_list()
        #辞書にする
        dict_of_coded_words[n] = temp_l
    #KHcoderのコーディングファイル形式にする
    list_for_write = []
    for n in l_c_n:
        #インデックス行をつくる
        index = "＊{}\n".format(dict_of_codename[n])
        list_for_write.append(index)
        #単語の行を作る
        temp_l_w = list(dict_of_coded_words[n])
        for w in temp_l_w[:-1:]:
            word = "{} or ".format(w)
            list_for_write.append(word)
        list_for_write.append("{}\n\n".format(temp_l_w[-1]))
    return(list_for_write)


#入力値をチェックするモジュール
def check_num(list_of_x, nt):
    n = len(list_of_x)
    while  True:
        flag = False
        for i in range(n):
            print(str(i) + ": " + str(list_of_x[i]))
        c = input("使用する{}をひとつ選んでください（数値で入力) :".format(nt))
        # 誤入力チェック
        # 入力値は数字か
        try:
            c = int(c)
        except ValueError as e:
            print("数値を入力してください。")
            flag = True
            continue
        #入力値は範囲内か
        if c > n-1:
            flag = True
            print("数値が間違っています。")
        if flag == False:
            break
    return(c)

#二つ目のシートのパスを確定する
def second_sheet(path):
    print("■コードの情報は同じファイルに入っていますか？")
    c = yes_or_no()
    if c == True:
        path = path
    else:
        typ = [("Excell", "xlsx")]
        path = tkinter.filedialog.askopenfilename(initialdir="C:/Users/lemat/Documents/ShineDoc/sources/",
                                                  filetypes=typ)
    return(path)


#シート名、列名を選択させる
def make_df(path, number, word):
    print("{}と{}を対応させます。".format(word, number))
    #シート名をリストにする
    list_of_sheets = list(xlrd.open_workbook(path).sheet_names())
    print("ファイル名とシート名は以下の通りです。")
    print(path)
    # シートを確定する
    nt = "シート"
    sh = check_num(list_of_sheets, nt)
    print("選択された{}は「{}」です。".format(nt, list_of_sheets[sh]))
    #列名を確定する
    df = pandas.read_excel(path,sheet_name=sh, index_col=None)
    # インデックス行をリストにする
    w = word
    nc = "列"
    lc = list(df.columns)
    # 単語の列を確定する
    print("■{}が含まれる{}を確定します。".format(w,nc))
    col = check_num(lc, nc)
    w_col = lc[col]
    print("選択された{}は「{}」です。".format(nc, w_col))
    # クラスタ番号の列を確定する
    w = number
    print("■{}が含まれる{}を確定します。".format(w, nc))
    col = check_num(lc, nc)
    n_col = lc[col]
    print("選択された{}は「{}」です。".format(nc, n_col))
    #選択した列でDFを作る
    df1 = df.loc[:, [n_col, w_col]]
    return(df1)


#ユーザーが指定するファイルを読み込む
def get_path():
    typ = [("Excell", "xlsx")]
    path = tkinter.filedialog.askopenfilename(initialdir="C:/Users/lemat/Documents/ShineDoc/sources/", filetypes=typ)
    return (path)


if __name__ == "__main__":
    path = get_path()
    df_code = make_df(path, "クラスタ番号", "抽出語")
    path = second_sheet(path)
    df_coding = make_df(path, "コード番号", "コード名")
    list_of_string = make_data(df_code, df_coding)
    print("■ファイルを保存します。")
    path = chose_path(path)
    write_file(list_of_string, path)