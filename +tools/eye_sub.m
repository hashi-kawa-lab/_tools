function out = eye_sub(A, s)
arguments
    A 
    s = 1
end

I = eye(size(A));
out = s*I - A;

end

