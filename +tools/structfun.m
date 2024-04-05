function Sout = structfun(func, S)

    fn = fieldnames(S);
    Sout = struct();

    for i = 1:numel(fn)
        Sout.(fn{i}) = func(S.(fn{i}));
    end

end