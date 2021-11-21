
## volume

このフォルダ配下には解析した結果のXMLが格納されます。


## template.xml

- child
    プロジェクトを構成する二元素の一つです。
    主に「定義」を表します。
    また、childの中にはchildを複数定義することができ、これはchild同志の包含,並列関係を表します。

    - kind
        要素の属性です。

        - class
            クラスのことです。
        
        - file
            ファイルのことです。

    - title
        要素の名前です。

    - description
        要素の説明です。

    - child
        内包する「定義」を表します。
        （folderにとってのfile,
        fileにとってのclass,
        classにとってのmethod,
        methodにとってのロジックを表します。)

- call
    プロジェクトを構成する二元素の一つです。
    主に「利用」を表します。

    - title 
        利用対象の名前を表します。

    - location
        利用対象の定義場所を表します。


