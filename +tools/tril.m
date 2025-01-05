function A = tril(A11, A21, A22)

A = [A11, zeros(size(A11, 1), size(A22, 2));
    A21, A22];

end

