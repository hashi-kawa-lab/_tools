function out = add_state_output_port(sys)


[ny, nu] = size(sys);
nx = order(sys);

[A, B, C, D] = ssdata(sys);

Cnew = [C; eye(nx)];
Dnew = blkdiag(D, zeros(nx, 0));

out = ss(A, B, Cnew, Dnew, sys.ts);
out.InputGroup = sys.InputGroup;
out.OutputGroup = sys.OutputGroup;

out.OutputGroup.x = size(C, 1) + (1:nx);


end