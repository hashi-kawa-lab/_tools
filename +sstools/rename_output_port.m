function sys = rename_output_port(sys, from, to)

if isfield(sys.OutputGroup, from)
    sys.OutputGroup.(to) = sys.OutputGroup.(from);
    sys.OutputGroup = rmfield(sys.OutputGroup, from);
end

end