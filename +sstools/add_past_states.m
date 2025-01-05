function [sys_out, T1, T2] = add_past_states(sys, n, C)

if nargin < 3
    C = eye(order(sys));
end

A_add = blkdiag(C, eye((n-1)*size(C, 1)));

Anew = [sys.a, zeros(size(sys.a, 1), (n-1)*size(C, 1)); A_add];
Anew = [Anew,  zeros(size(Anew, 1), size(C, 1))];
Cnew = [sys.C, zeros(size(sys.c, 1), n*size(C, 1))];
Bnew = [sys.b; zeros(size(A_add, 1), size(sys.b, 2))];
sys_out = ss(Anew, Bnew, Cnew, sys.d, sys.Ts);

T1 = tools.add_zeros(eye(order(sys)), 2, like=A_add, position="end");
T2 = tools.add_zeros(eye(size(A_add, 1)), 2, like=sys.a, position="start");

end

