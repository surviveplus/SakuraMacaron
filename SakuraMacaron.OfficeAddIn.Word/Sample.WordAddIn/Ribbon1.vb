﻿Imports Microsoft.Office.Tools.Ribbon
Imports Net.Surviveplus.SakuraMacaron.OfficeAddIn.UI
Imports Net.Surviveplus.SakuraMacaron.OfficeAddIn.Word

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        If Me.samplePane Is Nothing Then

            Dim c = New UserControl1
            AddHandler c.DoButtonClick,
                Sub(sender2, e2)

                    'ThisAddIn.Current.Application.Selection.Text = "sample"

                    Dim macaron = New WordMacaron(ThisAddIn.Current.Application)
                    macaron.ReplaceSelectionParagraphs(
                        Sub(a)
                            a.IsCanceled = a.Text.Contains("X")
                        End Sub,
                        Sub(a)
                            a.IsSkipped = a.Text = String.Empty OrElse (a.Text.StartsWith("[[") AndAlso a.Text.EndsWith("]]"))
                            a.InsertBeforeText = "[[ "
                            a.InsertAfterText = " ]]"
                        End Sub
                    )

                    macaron.ReplaceSelectionText(
                        Sub(a)
                            a.IsCanceled = a.Text.Contains("X")
                        End Sub,
                        Sub(a)
                            a.IsSkipped = a.Text = String.Empty OrElse (a.Text.StartsWith("START") AndAlso a.Text.EndsWith("END"))
                            a.InsertBeforeText = "START" & vbCrLf
                            a.InsertAfterText = vbCrLf & "END"
                        End Sub
                    )

                End Sub

            Me.samplePane = New ElementControlPane(Of UserControl1)(c)
            Me.samplePane.Pane = ThisAddIn.Current.CustomTaskPanes.Add(Me.samplePane.Control, "Sample Pane", ThisAddIn.Current.Application.ActiveWindow)
            Me.samplePane.Pane.Width = 350

        End If

        Me.samplePane?.Show()
    End Sub

    Private samplePane As ElementControlPane(Of UserControl1)


    Private kana As String = "ァアィイゥウヴェエォオヵカガキギキ゚ㇰクグク゚ヶケゲケ゚コゴコ゚サザㇱシジㇲスズセゼセ゚ソゾタダチヂッツヅツ゚テデㇳトドト゚ナニㇴヌネノㇵハバパㇶヒビピㇷフブㇷ゚プㇸヘベペㇹホボポマミㇺムメモャヤュユョヨㇻラㇼリㇽルㇾレㇿロヮワヷヰヸヱヹヲヺン"
    Private toLong As String() = {
        "アクセサ", "アクタ", "アクティベータ", "アグリゲータ", "アセンブラ", "アダプタ", "アップデータ", "アップローダ", "アドバイザ", "アドミニストレータ", "アナライザ", "アニメータ", "アロケータ", "アンインストーラ", "アンカ", "アンギュラ", "アンパッカ", "イコライザ", "イニシエータ", "イニシャライザ", "イネーブラ", "イメージセッタ", "インサータ", "インジェクタ", "インジケータ", "インストーラ", "インスペクタ", "インターセプタ", "インタープリタ", "インデクサ", "インテグレータ", "インバータ", "インバリデータ", "インポータ", "ウォーカ", "ウォータ", "ウォリア", "エキスパンダ", "エクスチェンジャ", "エクステンダ", "エクスプローラ", "エクスポータ", "エスカレータ", "エゼクタ", "エディタ", "エバリュエータ", "エミュレータ", "エレベータ", "エンコーダ", "オーガナイザ", "オートダイヤラ", "オートフィルタ", "オシレータ", "オフィサ", "オブザーバ", "オプティマイザ", "オペレータ", "オルタネータ", "カウンタ", "カスタマ", "カスタマイザ", "カテゴライザ", "ガバナ", "カプラ", "カムコーダ", "キャラクタ", "キュイラシェ", "クラスタ", "クリーナ", "クリエータ", "グリッパ", "クローラ", "コーディネータ", "コレクタ", "コンシューマ", "コンストラクタ", "コンセントレータ", "コンダクタ", "コンテナ", "コンデンサ", "コントリビュータ", "コントローラ", "コンバータ", "コンピュータ", "コンフィギュレータ", "コンプレッサ", "コンベクタ", "コンペンセータ", "コンポーザ", "コンポジタ", "サブウーファ", "サブコンテナ", "サブスクライバ", "サブフィルタ", "サブフォルダ", "サブマスタ", "サプライヤ", "サブレイヤ", "サプレッサ", "サマライザ", "シアタ", "シーケンサ", "シェーダ", "ジェネレータ", "シフタ", "シャッタ", "ジャンパ", "シリアライザ", "シリンダ", "シンクロナイザ", "シンセサイザ", "スイッチャ", "スーパーバイザ", "スーパバイザ", "スーペリア", "スカベンジャ", "スカラ", "スキャッタ", "スキャナ", "スクリーナ", "スクリーンセーバ", "スクリプタ", "スクレーパ", "スタッカ", "ステージャ", "ステマ", "ストーカ", "スニファ", "スパイダ", "スピア", "スピーカ", "スプーラ", "スプリッタ", "スペーサ", "スライサ", "スライダ", "セクタ", "セパレータ", "セレクタ", "センサ", "センタ", "センダ", "タイプライタ", "タイマ", "ダイヤラ", "ダウンローダ", "チェンジャ", "チャプタ", "チャレンジャ", "チューナ", "ディザ", "ディストリビュータ", "ディスパッチャ", "ディスペンサ", "ディバイダ", "ディフューザ", "ディレクタ", "テクスチャライザ", "デコーダ", "デザイナ", "デジタイザ", "デシリアライザ", "テスタ", "デストラクタ", "デスプーラ", "デバイザ", "デバッガ", "デベロッパ", "デマルチプレクサ", "デリミタ", "ドライバ", "ドライヤ", "トラクタ", "トラッカ", "トランシーバ", "トランスデューサ", "トランスファ", "トランスフォーマ", "トランスポンダ", "トランスレータ", "トリガ", "トリマ", "トレーサ", "トレーラ", "ドロア", "ナビゲータ", "ナレータ", "ノーマライザ", "パートナ", "バーバライザ", "ハイパーバイザ", "バインダ", "パケッタイザ", "パスファインダ", "バスマスタ", "パッケージャ", "バッファ", "パブリッシャ", "パラメータ", "バランサ", "バリオメータ", "バリデータ", "パルベライザ", "ハンドラ", "ハンマ", "ビジュアライザ", "ピッカ", "ビヘイビア", "ピュアライザ", "ビューア", "ビルダ", "ファイナライザ", "ファイバ", "ファインダ", "ファクタ", "フィーダ", "フィニッシャ", "フィルタ", "フィンガ", "ブースタ", "ブートローダ", "フェイダ", "フォルダ", "フォワーダ", "フッタ", "ブラウザ", "プランジャ", "プランナ", "プリマスタ", "プリンタ", "プルーナ", "ブレーカ", "プレースホルダ", "フレーバ", "プレゼンタ", "プレビューア", "プレフィルタ", "ブレンダ", "ブローカ", "ブロードキャスタ", "プロジェクタ", "ブロッカ", "プロッタ", "プロテクタ", "プロデューサ", "プロバイダ", "プロファイラ", "ベアラ", "ページャ", "ベクタ", "ヘッダ", "ヘリコプタ", "ヘルパ", "ベンダ", "ボイジャ", "ポインタ", "ポストマスタ", "ポリマ", "ホルダ", "マーシャラ", "マイナ", "マインスイーパ", "マスタ", "マッパ", "マニピュレータ", "マネージャ", "マルチフィーダ", "マルチプライヤ", "マルチプレクサ", "マルチマスタ", "マルチモニタ", "マルチレイヤ", "ミキサ", "ミニドライバ", "ミニフィルタ", "ミューテータ", "メッセンジャ", "メディエイタ", "メンバ", "メンバシップ", "モジュラ", "モディファイヤ", "モデレータ", "モニカ", "モニタ", "ライザ", "ライセンサ", "ライタ", "ラスタ", "ラスタライザ", "リアクタ", "リクエスタ", "リスナ", "リゾルバ", "リダイレクタ", "リパッケージャ", "リパブリッシャ", "リファクタ", "リフレッシャ", "リマインダ", "リンカ", "レイヤ", "レコーダ", "レコグナイザ", "レシーバ", "レジストラ", "レスポンダ", "レビューア", "レプリケータ", "レベラ", "レポータ", "レンダ", "レンダラ", "ローカライザ", "ロケータ",
        "アーチャ", "アウタ", "アウトロ", "アカデミ", "アスキ", "アッパ", "アドベンチャ", "アニュバ", "アバタ", "アバランチャ", "アフタ", "アベンジャ", "アレルギ", "アングラ", "アンサ", "アンダ", "アンダーバ", "イージ", "イースタ", "イェーガ", "イグナイタ", "インストラクタ", "インタビュ", "インナ", "インベンタ", "ウィスキ", "ウィスパ", "ウィッカ", "ウォーマ", "ウォッチャ", "エコ", "エコノミ", "エナジ", "エネルギ", "エラ", "エルダ", "エンタ", "エンチャンタ", "エンフォーサ", "オーサ", "オーダ", "オーナ", "オーバ", "オファ", "オリバ", "ガードナ", "カウンセラ", "カッタ", "カヌ", "カバ", "カラ", "カレ", "カレンダ", "カロリ", "カンガル", "ガンナ", "カンパニ", "キ", "ギタ", "キッカ", "キャスタ", "キャッチャ", "キャノニア", "キャブレタ", "ギャラリ", "キャンカ", "キュ", "キュ", "キラ", "クーラ", "グライダ", "クラッカ", "クラッシャ", "クランベリ", "クリーバ", "クル", "クルーザ", "クルセイダ", "グレ", "クレータ", "グレータ", "グロ", "クローバ", "クロスバ", "クロバ", "ゲートキーパ", "ゲーマ", "コーナ", "コーヒ", "コーラ", "ゴールキーパ", "コピ", "コマンダ", "コミュータ", "コンピテンシ", "サーバ", "サイドバ", "サイバ", "サッカ", "ザッパ", "サブキ", "サブツリ", "サマリ", "サラマンダ", "シーソ", "シーナリ", "シグナラ", "シチュ", "ジッタ", "シナジ", "シミタ", "ジャガ", "シャワ", "シャンプ", "ジュエリ", "シュガ", "ショ", "ジョーカ", "ショッカ", "シルバ", "スーパ", "スカーミッシャ", "スキ", "スキッタ", "スキナ", "スクーナ", "スクリュ", "スクロールバ", "スコーチャ", "スタ", "スタータ", "スターフラワ", "スティンガ", "ステッカ", "ストーリ", "ストッパ", "ストライダ", "ストロベリ", "スナイパ", "スニーカ", "スノ", "スパイカ", "スパッタ", "スピッタ", "スプリンクラ", "スプレ", "スポイラ", "スポンサ", "スラッシャ", "スリーパ", "スリンガ", "スル", "スロ", "スロータ", "セータ", "セーバ", "セッタ", "セプタ", "セミナ", "セルラ", "セレモニ", "ソータ", "ソーラ", "ソルジャ", "ソルバ", "タイガ", "ダイバ", "タイムリ", "ダガ", "タクシ", "ダミ", "タワ", "タンカ", "ダンサ", "ダンパ", "チータ", "チェッカ", "チェリ", "チャータ", "チャウダ", "チャネラ", "チョッパ", "チンパンジ", "ツア", "ツリ", "ティ", "ディーラ", "ディスカバ", "ディナ", "ディフェンダ", "デイリ", "ティンバ", "デスクバ", "デストロイヤ", "デスポイラ", "テレポータ", "テンキ", "ドゥエラ", "ドクタ", "トッパ", "ドップラ", "ドッペルゲンガ", "トナ", "トラベラ", "ドラマ", "ドリルスル", "ドルビ", "トレーダ", "トレーナ", "トレジャ", "ドレッサ", "トロフィ", "トロリ", "ナンバ", "ニードラ", "ニュ", "ニュースリーダ", "ニュースレタ", "ネイチャ", "ネービ", "バ", "パーサ", "バーサーカ", "バースデ", "バーテンダ", "バーナ", "ハーバ", "バーベキュ", "ハーモニ", "ハイカ", "ハイカラ", "バイザ", "ハイパ", "パイパ", "バイヤ", "パウダ", "ハウツ", "ハウラ", "バザ", "パスキ", "パススル", "バスタ", "バタ", "ハッカ", "バックラ", "バッシャ", "ハッピ", "パトローラ", "バナ", "パブリシティ", "バリュ", "バルコニ", "パワ", "ハンガ", "バンカ", "パンサ", "パンジ", "ハンタ", "パンチャ", "バンパ", "ハンバーガ", "ピアサ", "ヒータ", "ビーバ", "ヒーロ", "ビギナ", "ビクトリ", "ビジ", "ビジタ", "ピト", "ビフォ", "ビュ", "ヒューザ", "ピューリファイア", "ファイタ", "ファンシ", "ファンタジ", "フィッシャ", "フェールオーバ", "フェザ", "フェリ", "フォッカ", "フォトグラフィ", "フォロ", "ブザ", "ブッチャ", "フライヤ", "ブラスタ", "フラワ", "フリ", "フリークエンシ", "フリッカ", "フリッパ", "ブル", "ブルーベリ", "ブルドーザ", "プレイナ", "プレーヤ", "プレーリ", "プレオーダ", "プレデタ", "プレビュ", "フロ", "フロッピ", "プロンプタ", "ペイズリ", "ペーパ", "ベストセラ", "ペッカリ", "ペッパ", "ヘビ", "ベビ", "ベビーシッタ", "ベリ", "ヘルシ", "ベンチャ", "ボイラ", "ボ", "ボーダ", "ポータ", "ポスタ", "ホッケ", "ホットキ", "ホッパ", "ポニ", "ホバ", "ホビ", "ポピュラ", "ホラ", "ポリシ", "マーカ", "マーキ", "マッシャ", "マッチャ", "マナ", "マネ", "マホガニ", "ミステリ", "ミラ", "ムーバ", "ムービ", "メーカ", "メータ", "メーラ", "メジャ", "メニュ", "モータ", "モット", "モンスタ", "ユーザ", "ライダ", "ラグジュアリ", "ラズベリ", "ラダ", "ラッキ", "ラッシャ", "ラッパ", "ラプタ", "ラベンダ", "ランガ", "ランチャ", "ランデブ", "ランナ", "リーダ", "リセラ", "リビーラ", "リフラクタ", "リベレータ", "リレ", "ルータ", "ルーラ", "ルビ", "レーザ", "レーダ", "レギュラ", "レクサ", "レザ", "レジャ", "レスキュ", "レタ", "レッカ", "レッサ", "レトリバ", "レバ", "レビュ", "レンジャ", "ロータ", "ローダ", "ロータリ", "ローラ", "ロガ", "ロッカ", "ロビ", "ワーカ", "ワークフロ", "ワインセラ", "ワッカ"}
    Private toShort As String() = {"アウトドア", "アクセラレータ", "インテリア", "インドア", "エクステリア", "エンジニア", "ギア", "キャリア", "クリア", "コネクタ", "コンパイラ", "コンベヤ", "シニア", "ジュニア", "スケジューラ", "ステラ", "スリッパ", "センチメートル", "ターミネータ", "タール", "ドア", "トランジスタ", "ドル", "バザール", "バリア", "ピア", "ビール", "フォーマッタ", "プレミア", "フロア", "プログラマ", "プロセッサ", "プロペラ", "フロンティア", "ベア", "ボランティア", "ポリエステル", "ミリメートル", "メートル", "ユーモア", "ラジエータ", "リア", "リニア", "レジスタ"}
    Private mappingKanaKanji As String(,) = {{"挨拶", "あいさつ"}, {"曖昧", "あいまい"}, {"敢えて", "あえて"}, {"辺り", "あたり"}, {"あて名", "宛名"}, {"跡始末", "後始末"}, {"貴方", "あなた"}, {"貴女", "あなた"}, {"彼方", "あなた"}, {"あぶない", "危ない"}, {"余り", "あまり"}, {"予め", "あらかじめ"}, {"凡ゆる", "あらゆる"}, {"有難う", "ありがとう"}, {"在る", "ある"}, {"有る", "ある"}, {"或", "ある"}, {"言い替え", "言い換え"}, {"如何なる", "いかなる"}, {"幾つ", "いくつ"}, {"幾分", "いくぶん"}, {"幾ら", "いくら"}, {"幾らか", "いくらか"}, {"不可い", "いけない"}, {"椅子", "いす"}, {"致し", "いたし"}, {"戴く", "頂く"}, {"一斉", "いっせい"}, {"一層", "いっそう"}, {"一旦", "いったん"}, {"一杯", "いっぱい"}, {"未だに", "いまだに"}, {"愈々", "いよいよ"}, {"居る", "いる"}, {"色々", "いろいろ"}, {"所謂", "いわゆる"}, {"打合せ", "打ち合わせ"}, {"上手い", "うまい"}, {"巧い", "うまい"}, {"売上げ", "売り上げ"}, {"売りあげ", "売り上げ"}, {"売出し", "売り出し"}, {"売りだし", "売り出し"}, {"円錐", "円すい"}, {"於いて", "おいて"}, {"大凡", "おおよそ"}, {"置き替え", "置き換え"}, {"送り仮名", "送りがな"}, {"送りだす", "送り出す"}, {"於ける", "おける"}, {"恐らく", "おそらく"}, {"虞", "おそれ"}, {"各々", "おのおの"}, {"面白い", "おもしろい"}, {"凡そ", "およそ"}, {"及び", "および"}, {"甲斐", "かい"}, {"買い替え", "買い換え"}, {"買換え", "買い換え"}, {"却って", "かえって"}, {"係る", "かかる"}, {"掛かる", "かかる"}, {"係わり", "かかわり"}, {"関わり", "かかわり"}, {"係わる", "かかわる"}, {"関わる", "かかわる"}, {"書換え", "書き換え"}, {"書き替え", "書き換え"}, {"加減", "かげん"}, {"凝まる", "固まる"}, {"且つ", "かつ"}, {"括弧", "かっこ"}, {"嘗て", "かつて"}, {"仮名", "かな"}, {"可成り", "かなり"}, {"予て", "かねて"}, {"構わない", "かまわない"}, {"かも知れない", "かもしれない"}, {"硝子", "ガラス"}, {"可愛い", "かわいい"}, {"気がつく", "気が付く"}, {"規準", "基準"}, {"切っ掛け", "きっかけ"}, {"気づく", "気付く"}, {"切上げ", "切り上げ"}, {"切り換え", "切り替え"}, {"切替え", "切り替え"}, {"切下げ", "切り下げ"}, {"切捨て", "切り捨て"}, {"切取り", "切り取り"}, {"伐る", "切る"}, {"斬る", "切る"}, {"奇麗", "きれい"}, {"亀裂", "きれつ"}, {"極めて", "きわめて"}, {"究める", "極める"}, {"窮める", "極める"}, {"喰い違う", "食い違う"}, {"句切る", "区切る"}, {"下さい", "ください"}, {"括れる", "くびれる"}, {"凹み", "くぼみ"}, {"窪み", "くぼみ"}, {"汲む", "くむ"}, {"較べる", "比べる"}, {"繰上げ", "繰り上げ"}, {"繰下げ", "繰り下げ"}, {"委しい", "詳しい"}, {"後篇", "後編"}, {"小売", "小売り"}, {"極", "ごく"}, {"此処", "ここ"}, {"此所", "ここ"}, {"午后", "午後"}, {"箇数", "個数"}, {"擦る", "こする"}, {"此方", "こちら"}, {"是", "これ"}, {"此れ", "これ"}, {"此等", "これら"}, {"頃", "ごろ"}, {"差違", "差異"}, {"幸い", "さいわい"}, {"逆様", "逆さま"}, {"遡る", "さかのぼる"}, {"捜る", "探る"}, {"様々", "さまざま"}, {"更に", "さらに"}, {"更に", "さらに"}, {"仕上げる", "しあげる"}, {"強いて", "しいて"}, {"然し", "しかし"}, {"併し", "しかし"}, {"しかし乍ら", "しかしながら"}, {"而も", "しかも"}, {"仕組み", "しくみ"}, {"次第", "しだい"}, {"従って", "したがって"}, {"確り", "しっかり"}, {"屡々", "しばしば"}, {"暫く", "しばらく"}, {"仕舞う", "しまう"}, {"了う", "しまう"}, {"仕舞う", "しまう"}, {"受註", "受注"}, {"報らせる", "知らせる"}, {"知れない", "しれない"}, {"据えつける", "据え付ける"}, {"透き間", "すきま"}, {"隙間", "すきま"}, {"過ぎる", "すぎる"}, {"直ぐ", "すぐ"}, {"宛つ", "ずつ"}, {"棄てる", "捨てる"}, {"即ち", "すなわち"}, {"素早い", "すばやい"}, {"素早く", "すばやく"}, {"素晴らしい", "すばらしい"}, {"全て", "すべて"}, {"総て", "すべて"}, {"凡て", "すべて"}, {"為る", "する"}, {"摩れる", "擦れる"}, {"整頓", "整とん"}, {"是非", "ぜひ"}, {"副う", "そう"}, {"其処", "そこ"}, {"其処で", "そこで"}, {"害なう", "損なう"}, {"具える", "備える"}, {"其の上", "そのうえ"}, {"其の内", "そのうち"}, {"其の代わり", "その代わり"}, {"その為", "そのため"}, {"其の為", "そのため"}, {"其のまま", "そのまま"}, {"側", "そば"}, {"傍", "そば"}, {"其れ", "それ"}, {"夫々", "それぞれ"}, {"大した", "たいした"}, {"大して", "たいして"}, {"大体", "だいたい"}, {"大抵", "たいてい"}, {"大変", "たいへん"}, {"沢山", "たくさん"}, {"足す", "たす"}, {"援ける", "助ける"}, {"扶ける", "助ける"}, {"尋ねる", "たずねる"}, {"唯", "ただ"}, {"只", "ただ"}, {"叩く", "たたく"}, {"但し", "ただし"}, {"畳む", "たたむ"}, {"私達", "私たち"}, {"君達", "君たち"}, {"あなた達", "あなたたち"}, {"発つ", "たつ"}, {"経つ", "たつ"}, {"例え", "たとえ"}, {"喩え", "たとえ"}, {"仮令", "たとえ"}, {"例えば", "たとえば"}, {"例える", "たとえる"}, {"喩える", "たとえる"}, {"愉しい", "楽しい"}, {"たのむ", "頼む"}, {"度重なる", "たび重なる"}, {"度々", "たびたび"}, {"偶に", "たまに"}, {"溜まる", "たまる"}, {"為", "ため"}, {"為に", "ために"}, {"貯める", "ためる"}, {"溜める", "ためる"}, {"誰", "だれ"}, {"滴れる", "たれる"}, {"丁度", "ちょうど"}, {"就いて", "ついて"}, {"次いで", "ついで"}, {"序でに", "ついでに"}, {"就いては", "ついては"}, {"遂に", "ついに"}, {"掴む", "つかむ"}, {"点く", "つく"}, {"付く", "づく"}, {"作りあげる", "作り上げる"}, {"作りかえる", "作り替える"}, {"作りかた", "作り方"}, {"作りだす", "作り出す"}, {"作りなおす", "作り直す"}, {"附け替える", "付け替える"}, {"附け加える", "付け加える"}, {"付け足す", "付けたす"}, {"附け足す", "付けたす"}, {"勉めて", "つとめて"}, {"努めて", "つとめて"}, {"繋がる", "つながる"}, {"繋ぐ", "つなぐ"}, {"潰す", "つぶす"}, {"摘まむ、撮む", "つまむ"}, {"詰まり", "つまり"}, {"積もり", "つもり"}, {"剛い", "強い"}, {"辛い", "つらい"}, {"釣り合う", "つりあう"}, {"釣り下げる", "つり下げる"}, {"吊り下げる", "つり下げる"}, {"吊るす", "つるす"}, {"出合う", "出会う"}, {"丁寧", "ていねい"}, {"出来", "でき"}, {"出来上がり", "できあがり"}, {"出来上がる", "できあがる"}, {"適確", "的確"}, {"出来る", "できる"}, {"手だすけ", "手助け"}, {"手つづき", "手続き"}, {"で無い", "でない"}, {"手なおし", "手直し"}, {"手許", "手元"}, {"照らしあわせ", "照らし合わせ"}, {"問いあわせ", "問い合わせ"}, {"と言う", "という"}, {"と云う", "という"}, {"同士", "どうし"}, {"通り", "とおり"}, {"時々", "ときどき"}, {"何処", "どこ"}, {"所が", "ところが"}, {"所で", "ところで"}, {"処で", "ところで"}, {"綴じる", "とじる"}, {"何方", "どちら"}, {"止まる", "とどまる"}, {"止める", "とどめる"}, {"留める", "とどめる"}, {"何の様な", "どのような"}, {"取り敢えず", "とりあえず"}, {"取りあげる", "取り上げる"}, {"取扱い", "取り扱い"}, {"取りあつかう", "取り扱う"}, {"取りいれ", "取り入れ"}, {"取りいれる", "取り入れる"}, {"取り換え、取替え", "取り替え"}, {"取り換える", "取り替える"}, {"取決め", "取り決め"}, {"取り極める", "取り決める"}, {"取消し", "取り消し"}, {"取りけす", "取り消す"}, {"取りこむ", "取り込む"}, {"取りさる", "取り去る"}, {"取りそろえる", "取り揃える"}, {"取りだす", "取り出す"}, {"取付け", "取り付け"}, {"取りつける", "取り付ける"}, {"取りのこす", "取り残す"}, {"取りはらう", "取り払う"}, {"採る", "とる"}, {"録る", "とる"}, {"盗る", "とる"}, {"摂る", "とる"}, {"何れ", "どれ"}, {"無い", "ない"}, {"尚", "なお"}, {"中々", "なかなか"}, {"却々", "なかなか"}, {"乍ら", "ながら"}, {"無くす", "なくす"}, {"無くなる", "なくなる"}, {"無し", "なし"}, {"何故", "なぜ"}, {"なつ印", "捺印"}, {"等", "など"}, {"倣う", "ならう"}, {"並びに", "ならびに"}, {"成る", "なる"}, {"成る可く", "なるべく"}, {"馴れる", "慣れる"}, {"何ら", "なんら"}, {"難い", "にくい"}, {"に過ぎない", "にすぎない"}, {"贋物", "偽物"}, {"に就いて", "について"}, {"に就き", "につき"}, {"に連れて", "につれて"}, {"に成る", "になる"}, {"にも拘らず", "にもかかわらず"}, {"に因って", "によって"}, {"狙う", "ねらう"}, {"乗換え", "乗り換え"}, {"ハガキ", "はがき"}, {"葉書", "はがき"}, {"捗る", "はかどる"}, {"許り", "ばかり"}, {"莫大", "ばく大"}, {"筈", "はず"}, {"果たして", "はたして"}, {"填め込む", "はめ込む"}, {"填める", "はめる"}, {"張り付ける", "貼り付ける"}, {"張る", "貼る"}, {"引換え", "引き換え"}, {"引き替え", "引き換え"}, {"引き替える", "引き換える"}, {"引きつづく", "引き続く"}, {"引きぬく", "引き抜く"}, {"引きのばす", "引き伸ばす"}, {"引きのばす", "引き延ばす"}, {"日毎", "日ごと"}, {"日頃", "日ごろ"}, {"歪む", "ひずむ"}, {"一通り", "ひととおり"}, {"一先ず", "ひとまず"}, {"一目", "ひとめ"}, {"日日", "日にち"}, {"拡がる", "広がる"}, {"拡げる", "広げる"}, {"附加", "付加"}, {"拭き取る", "ふき取る"}, {"脹れる", "膨れる"}, {"塞ぐ", "ふさぐ"}, {"相応しい", "ふさわしい"}, {"附属", "付属"}, {"蓋", "ふた"}, {"附着", "付着"}, {"振り仮名", "ふりがな"}, {"附録", "付録"}, {"併行", "並行"}, {"頁", "ページ"}, {"放る", "ほうる"}, {"殆ど", "ほとんど"}, {"前書き", "まえがき"}, {"委せる", "任せる"}, {"捲く", "巻く"}, {"正に", "まさに"}, {"先ず", "まず"}, {"益々", "ますます"}, {"又", "また"}, {"未だ", "まだ"}, {"又は", "または"}, {"全く", "まったく"}, {"纏める", "まとめる"}, {"真似", "まね"}, {"眩しい", "まぶしい"}, {"侭", "まま"}, {"儘", "まま"}, {"磨耗", "摩耗"}, {"稀", "まれ"}, {"希", "まれ"}, {"廻す", "回す"}, {"見出す", "見いだす"}, {"研く", "磨く"}, {"見付ける", "見つける"}, {"見積り", "見積もり"}, {"見做す", "見なす"}, {"見映え", "見栄え"}, {"向きあう", "向き合う"}, {"寧ろ", "むしろ"}, {"無闇に", "むやみに"}, {"無暗に", "むやみに"}, {"メガネ", "眼鏡"}, {"滅多", "めった"}, {"申しあげる", "申し上げる"}, {"申しこみ", "申し込み"}, {"申し込み書", "申込書"}, {"申しこむ", "申し込む"}, {"申しわけ", "申し訳"}, {"若し", "もし"}, {"若しくは", "もしくは"}, {"持ちはこぶ", "持ち運ぶ"}, {"勿論", "もちろん"}, {"元々", "もともと"}, {"貰う", "もらう"}, {"洩れる", "漏れる"}, {"易しい", "やさしい"}, {"易い", "やすい"}, {"矢張り", "やはり"}, {"やむをえず", "やむを得ず"}, {"止める", "やめる"}, {"稍", "やや"}, {"遣り直す", "やり直す"}, {"軟らかい", "やわらかい"}, {"故に", "ゆえに"}, {"歪む", "ゆがむ"}, {"様子", "ようす"}, {"様な", "ような"}, {"様に", "ように"}, {"漸く", "ようやく"}, {"良く", "よく"}, {"善く", "よく"}, {"能く", "よく"}, {"余程", "よほど"}, {"拠り所", "よりどころ"}, {"因る", "よる"}, {"依る", "よる"}, {"宜しく", "よろしく"}, {"立派", "りっぱ"}, {"連繋", "連係"}, {"分かる", "わかる"}, {"解る", "わかる"}, {"判る", "わかる"}, {"という訳で", "というわけで"}, {"僅か", "わずか"}, {"亘って", "わたって"}, {"亘る", "わたる"}, {"割当て", "割り当て"}, {"割に", "わりに"}, {"身につけ", "身に付け"}}

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click

        Dim macaron = New WordMacaron(ThisAddIn.Current.Application)
        macaron.ReplaceSelectionParagraphs(
            Nothing,
            Sub(a)
                For Each item As String In toLong
                    Dim reg As New RegularExpressions.Regex("(" & item & ")([^ー" & kana & "]|$)")
                    a.Text = reg.Replace(a.Text, "$1ー$2")
                Next item

                a.Text = a.Text.Replace("エンタプライズ", "エンタープライズ")

                For Each item As String In toShort
                    Dim reg As New RegularExpressions.Regex("(" & item & ")(ー)")
                    a.Text = reg.Replace(a.Text, "$1")
                Next item

                For index = 0 To UBound(mappingKanaKanji) - 1
                    Dim s = mappingKanaKanji(index, 0)
                    Dim d = mappingKanaKanji(index, 1)
                    Dim reg As New RegularExpressions.Regex(s)
                    a.Text = reg.Replace(a.Text, d)
                Next
            End Sub)

    End Sub
End Class
