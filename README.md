# Otoha0922
(training on September 22)
(added on October 4)

### この作品について *（2022/10/4現在）*
本コードは、某イカゲームの試合に関する記載をExcelファイルに書き込みをすることを目的としています。
そのため、ゲームをプレイしていない場合、リスト内の単語に不明な点があると思いますがご了承ください。

### コードの概要 *（2022/10/4現在）*
本コードは、試合のステージとルール、試合の勝敗をプルダウン式で選択して、splatoon3というExcelファイルのA～C列に空白の行から記載するという流れになっています。
試合のステージとルールは、*2022/10/4*の時点で行えるものとなっています。ステージは全ステージを、ルールはプライべートマッチを除いた通常時にPvP（人対人）で行うことのできるものの略称をリストにしています。
また、勝敗は〇が勝利を、×が敗北を意味しております。
なお、今後勝率の算出や削除機能も実装を検討しております。

### 事前準備 *（2022/10/4現在）*
本コードの使用にはsplatoon3というExcelファイルが必要です。

### 操作手順 *（2022/10/4現在）*
プログラムを動かしたら、ステージを表示されるウィンドウのプルダウンから選択し、_OK_ をおしてください。もし、キャンセルをする場合は、_cansel_ を押してください。同様の手順でルール、勝敗を選択してください。3つそれぞれ選択が終わるとsplatoon3に書き込みが行われます。
