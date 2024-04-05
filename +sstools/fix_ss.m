function sys_ = fix_ss(sys)
%% InputGroupとOutputGroupに入っていない入出力を削除
%% 1つの入出力に複数の名前がついていた場合に複数の入出力に分ける
%% InputGroup.v1: 1, InputGroup.v2:1 (1入力)
%% ⇒
%% InputGroup.v1: 1, InputGroup.v2:2 (2入力)

fields = fieldnames(sys.OutputGroup);
C = cell(numel(fields), 1);
D = cell(numel(fields), 1);
og = struct();
idx = 0;
for i = 1:numel(fields)
    [~, ~, C{i}, D{i}] = ssdata(sys(fields{i}, :));
    og.(fields{i}) = idx + (1:size(C{i}, 1));
    idx = idx + size(C{i}, 1);
end

sys_ = ss(sys.a, sys.b, vertcat(C{:}), vertcat(D{:}), sys.Ts);
sys_.InputGroup = sys.InputGroup;
sys_.OutputGroup = og;
sys_.StateName = sys.StateName;

sys = sys_;

fields = fieldnames(sys.InputGroup);
B = cell(numel(fields), 1);
D = cell(numel(fields), 1);
ig = struct();
idx = 0;
for i = 1:numel(fields)
    [~, B{i}, ~, D{i}] = ssdata(sys(:, fields{i}));
    ig.(fields{i}) = idx + (1:size(B{i}, 2));
    idx = idx + size(B{i}, 2);
end

sys_ = ss(sys.a, horzcat(B{:}), sys.c, horzcat(D{:}), sys.Ts);
sys_.OutputGroup = sys.OutputGroup;
sys_.InputGroup = ig;
sys_.StateName = sys.StateName;

end

