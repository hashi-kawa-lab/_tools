function toggle_lines(obj, line_indices)
% toggle_lines(obj, line_indices)
%   入力されたFigure (gcf) または Axes (gca) オブジェクト内の、
%   指定された複数のラインプロットの可視性を切り替えます。
%
%   入力:
%     obj: 対象となるFigureまたはAxesのハンドル (例: gcf, gca)
%     line_indices: 表示/非表示を切り替えるラインのインデックスのベクトル (例: [1, 3])

    % 入力オブジェクトのタイプを確認し、Axesオブジェクトを取得
    if strcmpi(get(obj, 'Type'), 'Figure')
        ax = findall(obj, 'Type', 'Axes');
    elseif strcmpi(get(obj, 'Type'), 'Axes')
        ax = obj;
    else
        error('Input object must be a Figure or Axes handle.');
    end
    
    if isempty(ax)
        error('No Axes object found in the input Figure.');
    end
    
    % Axes内のラインプロットオブジェクトを取得
    all_lines = flipud(findall(ax, 'Type', 'Line'));
    if (ischar(line_indices) || isstring(line_indices)) && strcmpi(line_indices, 'all')
        line_indices = 1:numel(all_lines);
    end
    % 各インデックスに対してループ処理を実行
    for idx = line_indices(:)'
        if idx > 0 && idx <= length(all_lines)
            target_line = all_lines(idx);
            current_visibility = get(target_line, 'Visible');
            
            if strcmp(current_visibility, 'on')
                set(target_line, 'Visible', 'off');
            else
                set(target_line, 'Visible', 'on');
            end
        else
            fprintf('Warning: Invalid line index %d. The axes contains %d lines.\n', idx, length(all_lines));
        end
    end
    
    drawnow;
end