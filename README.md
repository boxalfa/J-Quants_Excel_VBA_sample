# J-Quants_Excel_VBA_sample
J-Quants API のフリープランで利用可能な機能をExcel VBA で実装しました。

１）機能

    リフレッシュトークン取得(/token/auth_user)
    
    IDトークン取得(/token/auth_refresh)
    
    上場銘柄一覧(/listed/info)
    
    株価四本値(/prices/daily_quotes)
    
    財務情報(/fins/statements)
    
    決算発表予定日(/fins/announcement)
    
２）利用方法

    利用には、シート「token」の指定セルにメールアドレスとパスワードを入れてください。

    事前準備
    1.EXCELのメニューからファイル－＞オプション－＞リボンのユーザ設定で開発にチェックを入れる。結果メニューに「開発」が表示されます。
    2.開発－＞Visual Basic を選択、標準モジュールに上記「３．提供ファイル」、2,3 のモジュールを追加します。
    3.VBA のメニュー「ツールー＞参照設定」で以下２つをチェック（参照設定）します。
        ・Microsoft XML, v6.0
        ・Microsoft Scripting Runtime
    ※EXCEL を新規にインストールした状態で動作確認をしておりますので、EXECEL の各種設定等を変更されている場合は初期状態に戻しご利用下さい。

    また、VBA での JSON 文字列解析に VBA-JSON v2.3.1（https://github.com/VBA-tools/VBA-JSON/releases/tag/v2.3.1）（MIT License）を利用しています。


３）免責

このソフトウェアを使用したことによって生じたすべての障害・損害・不具合等に関して、私と私の関係者および私の所属するいかなる団体・組織とも、一切の責任を負いません。各自の責任においてご使用ください。

