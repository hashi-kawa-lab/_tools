function out = add_state_input_port(sys)


[ny, nu] = size(sys);
nx = order(sys);

[A, B, C, D] = ssdata(sys);

Bnew = [B, eye(nx)];
Dnew = blkdiag(D, zeros(0, nx));

out = ss(A, Bnew, C, Dnew, sys.ts);
out.InputGroup = sys.InputGroup;
out.OutputGroup = sys.OutputGroup;

out.InputGroup.x = size(B, 2) + (1:nx);


end