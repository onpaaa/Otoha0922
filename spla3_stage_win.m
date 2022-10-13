%　本コードは某海鮮ゲームのステージとそのルールの勝敗を記録することを目的としている

%　記入するxlsxファイル先の指定
filename = "spla3.xlsx";

% ステージのリスト
stagelist = {'ユノハナ大渓谷', 'ゴンズイ地区', 'ヤガラ市場', 
    'マテガイ放水路', 'ナメロウ金属', 'マサバ海峡大橋', 
    'キンメダイ美術館', 'マヒマヒリゾート＆スパ', '海女美術大学', 
    'チョウザメ漁船','ザトウマーケット', 'スメーシーワールド'};

% ルールのリスト
rulelist = {'ナワバリ', 'エリア', 'ヤグラ', 'ホコ', 'アサリ'};

% 勝敗のリスト
winloselist = {'勝ち', '負け'};


% ステージのリスト選択をダイアログボックスで表示
[stage,tfstage] = listdlg('PromptString', {'ステージを選択',...
    },...
    'SelectionMode', 'single', 'ListString', stagelist);
% OKを選択
if tfstage == 1;
   stage = stagelist(stage);

    % ルールのリスト選択をダイアログボックスで表示
    rulelist = {'ナワバリ', 'エリア', 'ヤグラ', 'ホコ', 'アサリ'};
    [rule,tfrule] = listdlg('PromptString', {'ルールを選択',...
             },...
    'SelectionMode', 'single', 'ListString', rulelist);
    % OKを選択
    if tfrule == 1;
        rule = rulelist(rule);
        
        % 勝敗のリスト選択をダイアログボックスで表示
        [winlose,tfwinlose] = listdlg('PromptString', {'勝敗を選択',...
           },...
           'SelectionMode', 'single', 'ListString', winloselist);
        % OKを選択
        if tfwinlose == 1; 
            winlose = winloselist(winlose);

            %　コマンドウィンドウでの確認用
            % stage
            % rule
            % winlose

            %入力した結果のスプレッドシートへの記入
            matchresult = horzcat(stage, rule, winlose);
            writecell(matchresult, filename, 'Sheet', 1, 'WriteMode', 'append');

            % コマンドウィンドウでの確認用
            % readtable(filename)
            
            % お疲れ様です。ダイアログボックスの表示やスプレッドシートへの記入方法など勉強になりました。
            % 特に新しい機能など思い浮かばなかったんですが、
            % 選択を終えた後にexeleに書き込みが実行されているか確認するために、
            % 自分でファイルを開かなければいけないところが面倒に感じました。
            % 書き込みが完了した旨を、何らかの方法でユーザーにわかるように示せると親切だと思いました。山本
            open(filename);
        end
    end
end
