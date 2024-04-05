function sys = remove_port(sys, name_port)

sys = sstools.remove_input_port(sys, name_port);
sys = sstools.remove_output_port(sys, name_port);

end

