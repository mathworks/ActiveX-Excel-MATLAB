% Copyright (c) 2020, The MathWorks, Inc.
function autoFitCellWidth(filename)

    % Excel ファイルへの絶対パスを取得
    filepath = which(filename);
    
    % Excel に対して ActiveX を開く
    h = actxserver('excel.application');
    wb = h.WorkBooks.Open(filepath);
    
    % UsedRange: データが入っている範囲の
    % EntireColumn: 列全体を
    % AutoFit: データに合わせた幅にします
    wb.ActiveSheet.UsedRange.EntireColumn.AutoFit;
    
    % 指定したファイル名で保存しエクセルを閉じる
    wb.SaveAs(filename);
    wb.Close;
    h.Quit;
    h.delete;
    % 注意：この辺キッチリ Close/Quit/delete しておかないとあとでややこしいです。
    %（ほかのアプリで使われていて開けない・消せないなど起こります）
    % PC 再起動すれば大丈夫です。
end