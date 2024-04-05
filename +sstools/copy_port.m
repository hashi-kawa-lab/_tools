function sys = copy_port(sys, from, to)

sys = sstools.copy_input_port(sys, from, to);
sys = sstools.copy_output_port(sys, from, to);

end

