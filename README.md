# ReadableChart(what is this?)

you can automatically build "call graph" from source code written in vba and vb6.

<img src="https://raw.githubusercontent.com/kawadasatoshi/ReadableChart/main/sample.svg">

this is from <a href="https://github.com/jpimbert/VBAToolKit/blob/master/Source/ConfProd/vtkReferenceManager.cls">VBAToolKit</a>



## how to install(インストール方法)

run next command in powershell terminal

<pre><code>
git clone git@github.com:kawadasatoshi/ReadableChart.git

</code></pre>

## how to run(実行方法)

run next command and open src.svg

<pre><code>
./main.ps1              #build src.xml and src.uml
plantuml -svg src.uml   #build src.svg

</code></pre>










## 本アプリのカバー範囲

"メタプログラミング"などの特殊なプログラミング技法を除く全てのアプリケーションを構成するコードのこと

ここでの"全体像"とは

- クラス設計図

- 関数レベルの設計図

など



## 実現する機能

- 解析機能

  - クラスレベル

  - 関数、メソッドレベル

の粒度でのXMLを作成すること

- 可視化機能

上記のXMLからUMLなどの形式で図表を作成しコードの把握を達成すること





## 言語について

解析用の言語はpowershellを使用。










