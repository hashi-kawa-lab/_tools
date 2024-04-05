function sys = rename_port(sys, from, to)

sys = sstools.rename_input_port(sys, from, to);
sys = sstools.rename_output_port(sys, from, to);

end