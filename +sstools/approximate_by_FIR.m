function out = approximate_by_FIR(sys, N)

imp = impulse(sys, (N-1)*sys.ts);

if size(sys, 2) == 1
Ts = sys.ts;   
A_fir = blkdiag(zeros(1, 0), eye(N-2), zeros(0, 1));
B_fir = zeros(size(A_fir, 1), 1);
B_fir(1) = 1;
C_fir = imp(2:end, :)'*Ts;
D_fir = imp(1, :)'*Ts;

out = ss(A_fir, B_fir, C_fir, D_fir, Ts);

else
    out = tools.harrayfun(@(i) sstools.approximate_by_FIR(sys(:, i), N), 1:size(sys, 2));    
end

out.InputGroup = sys.InputGroup;
out.OutputGroup = sys.OutputGroup;

end

