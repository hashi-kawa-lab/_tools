function sys = copy_input_port(sys, from, to)

sys.InputGroup.(to) = sys.InputGroup.(from);
sys = sstools.fix_ss(sys);

end

