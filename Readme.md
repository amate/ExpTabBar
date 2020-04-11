# ExpTabBar

エクスプローラーにタブバーを追加します

## 動作確認環境
- Windows 10 home 64bit バージョン 1909

おそらく Windows10 以外では動きません  
v1.11以前なら動くかも…

## 導入方法

ダウンロードしたzipファイルを適当なフォルダへ展開します  
※一度インストールすると、アンインストールするまでファイルの移動ができないのでファイルの置き場所に注意してください  

自分の使っているOSが64bit環境なら x64フォルダの install64.batを、32bit環境なら x86フォルダの install.batをダブルクリックして実行  
ユーザーアカウント制御のダイアログが出るので、よろしければ "はい" を選択  
"DllRegisterServerは成功しました。" と出ればインストール成功です

ここで一度、エクスプローラーの再起動またはOSの再起動をします　　
　　
その後、エクスプローラーを起動  
エクスプローラー上部の"表示"をクリック  
エクスプローラーの上部右に表示される "オプション"の**文字列部分**をクリック(アイコン部分じゃないよ！)  
表示されるメニューにある"ExpTabBar"を選択すれば、タブバーが表示されるはずです


## 免責
作者(原著者＆改変者)は、このソフトによって生じた如何なる損害にも、
修正や更新も、責任を負わないこととします。
使用にあたっては、自己責任でお願いします。
 
何かあれば下記のURLにあるメールフォームか、githubのIssueにお願いします。　　
https://ws.formzu.net/fgen/S37403840/

## 著作権表示
Copyright (C) 2020 amate

私が書いた部分のソースコードは、MIT License とします。

## ビルド方法
Visual Studio 2019 が必要です  
ビルドには boost(1.72~)とWTL(10_9163) が必要なのでそれぞれ用意してください。

Boost::Log Boost::serializationを使用しているので、事前にライブラリのビルドが必要になります

Boostライブラリのビルド方法
https://sites.google.com/site/boostjp/howtobuild

 <pre>
コマンドライン
// x86
b2.exe install --prefix=lib toolset=msvc-14.2 runtime-link=static --with-thread --with-date_time --with-timer --with-log --with-serialization
// x64
b2.exe install --prefix=lib64 toolset=msvc-14.2 runtime-link=static address-model=64 --with-thread --with-date_time --with-timer --with-log --with-serialization
</pre>

## 使用ライブラリ

- boost  
https://www.boost.org/

- WTL  
http://sourceforge.net/projects/wtl/

- MinHook  
http://www.codeproject.com/KB/winsdk/LibMinHook.aspx

- nlohmann json   
https://github.com/nlohmann/json



# 更新履歴

<pre>
v1.12
・[add] タブグループ機能を追加
・[add] ダークモードに対応した
・[add] APIフックを使うかどうかはオプションで選べるようにした
・[add] "同名のタブが存在する時、タブの親フォルダ名を同時に表示する" オプションを追加
・[add] エクスプローラーの一列選択表示の有無に関わらず、タブ切り替えでスクロール位置の復元ができるようになった

・[change] CreateProcessはAPIフックしないようにした
・[change] タブリストの保存形式をxmlからjsonへ変更(今バージョンでは自動的にjsonに変換して保存するが、次回のバージョンからは自動変換機能は消す予定)
・[change] 最近閉じたタブの保存形式をxmlからjsonへ変更(以下同文)
・[change] タブリストの保存は一時ファイルを経由して行うようにした(運が悪いとタブリストが消滅する事への対策)
(アイコンと詳細表示に限る)
・[change] リストビューにおいて、移動キー入力でいつもサムネイルツールチップが表示されていたのを、サムネイルツールチップが表示されている時だけ選択画像を切り替えて表示するようにした
・[change] インストールやアンインストール時にエクスプローラー関連のレジストリキーを変更するのをやめた
・[change] オプションのデフォルト設定値をいくつか変更

・[fix] リストビューの一列表示有効時、戻るボタン押しながらホイールスクロールで画像切り替えを実行した場合、マウスホバーでサムネイルツールチップが表示されなくなるバグを修正
・[fix] タブの大きさを指定するのチェックが外れている時、タブの高さの計算がおかしかったのを修正
・[fix] サムネイルツールチップを無効にしているのに、ホイールクリックでサムネイルツールチップが表示されてしまうのを修正

・[misc] 開発環境をVisualStudio2019へ移行
・[misc] コンパイラのwarningに対処した
・[misc] windows7関連のコードを削除


</pre>
