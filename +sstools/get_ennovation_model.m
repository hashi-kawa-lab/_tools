function [sys, Sigma, kalmf, kalmp] = get_ennovation_model(sys, Q)

[A, B, C, D] = ssdata(sys);
Ts = sys.Ts;

Q_ = [B; D] * Q * [B; D]';


ny = size(D, 1);
R = Q_(end-ny+1:end, end-ny+1:end);
Q = Q_(1:end-ny, 1:end-ny);
S = Q_(1:end-ny, end-ny+1:end);

% sys_ = ss(A, eye(size(A)), C, 0, []);

% [kalmf, L, P] = kalman(sys_, Q, R, S);

Abar = A-S/R*C;
Qbar = Q-S/R*S';

[P_,~,~,INFO] = idare(Abar',C',Qbar, R);
K = P_*C'/(C*P_*C'+R);

K_pred = A*K + S/R*tools.eye_sub(C*K);

kalmf = ss(A-K_pred*C, K_pred, tools.eye_sub(K*C), K);
kalmp = ss(A-K_pred*C, K_pred, C, 0);

Sigma = C*P_*C' + R;
sys = ss(A, K_pred, C, eye(size(K_pred, 2)), Ts);

end

