function out = add_zeros(A, axis, n, options)
arguments
    A (:,:)
    axis (1,1) {mustBeMember(axis, [1,2])}
    n (1,1) {mustBeInteger} = 0
    options.like
    options.position {mustBeMember(options.position, {'start', 'end'})} = 'end'
end

if isfield(options, 'like')
    n = size(options.like, axis);
end

if strcmp(options.position, 'start')
    if axis == 1
        out = [zeros(n, size(A,2)); A];
    else
        out = [zeros(size(A,1), n), A];
    end
else
    if axis == 1
        out = [A; zeros(n, size(A,2))];
    elseif axis == 2
        out = [A, zeros(size(A,1), n)];
    end
end

end

