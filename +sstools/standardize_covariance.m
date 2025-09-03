function [sys_new, L] = standardize_covariance(sys, name, Q)


% L = chol(Q, 'lower')';

[U, S] = svd(Q);
L = U*diag(sqrt(diag(S)))*U';

[~, B, ~, D] = ssdata(sys(:, name));

% 新しい行列を計算します。
B_new = B * L;
D_new = D * L;

if iscell(name)
    input_index = tools.hcellfun(@(n) sys.InputGroup.(n), name);
else
    input_index = sys.InputGroup.(name);
end

[A, B, C, D, Ts] = ssdata(sys);

B(:, input_index) = B_new;
D(:, input_index) = D_new;

sys_new = ss(A, B, C, D, Ts);
sys_new.InputGroup = sys.InputGroup;
sys_new.OutputGroup = sys.OutputGroup;

end