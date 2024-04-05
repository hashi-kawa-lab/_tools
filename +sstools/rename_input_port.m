function sys = rename_input_port(sys, from, to)

if isfield(sys.InputGroup, from)
    sys.InputGroup.(to) = sys.InputGroup.(from);
    sys.InputGroup = rmfield(sys.InputGroup, from);
end

end