function [sys_new] = add_output_port(sys, options)
arguments (Input)
    sys
    options.name
    options.C
    options.D
end

[A, B, C, D, Ts] = ssdata(sys);

if ~isfield(options, 'C') && ~isfield(options, 'D')
    error('Must specify at least one of C or D matrices for the new output port');
end

if ~isfield(options, 'C')
    options.C = zeros(size(options.D, 1), size(A, 2));
end
if ~isfield(options, 'D')
    options.D = zeros(size(options.C, 1), size(B, 2));
end

Cnew = [C; options.C];

Dnew = [D; options.D];

sys_new = ss(A, B, Cnew, Dnew, Ts);
sys_new.InputGroup = sys.InputGroup;
sys_new.OutputGroup = sys.OutputGroup;

if isfield(options, 'name')
    sys_new.OutputGroup.(options.name) = size(sys, 1) + (1:size(options.C, 1));
end

end