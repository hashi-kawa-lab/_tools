function sys = copy_output_port(sys, from, to)

sys.OutputGroup.(to) = sys.OutputGroup.(from);
sys = sstools.fix_ss(sys);

end

