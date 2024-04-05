function sys = connect_by_group(varargin)

n = nargin-1;
ig = cell(n, 1);
og = cell(n, 1);
ig_adds = struct();
og_adds = struct();

for k = 1:n
    sys = varargin{k};
    ig{k} = fieldnames(sys.InputGroup);
    og{k} = fieldnames(sys.OutputGroup);
end

ig_all = vertcat(ig{:});
og_all = vertcat(og{:});

ig_unique = unique(ig_all);
og_unique = unique(og_all);

sys_adds_ig = cell(numel(ig_unique), 1);
sys_adds_og = cell(numel(ig_unique), 1);

for i = 1:numel(ig_unique)
    ig_adds = struct();
    is_overwrap = false(n, 1);
    for j = 1:n
        if sum(strcmp(ig{j}, ig_unique{i})) ~= 0
            is_overwrap(j) = true;
        end
    end
    if sum(is_overwrap) ~= 1
        n_i = nan;
        sys_adds = cell(n, 1);
        idx_in = 0;
        for j = 1:n
            if is_overwrap(j)
                name = ig_unique{i};
                sys = varargin{j};
                name_new = strcat(name, '_input_', num2str(j));
                sys.InputGroup.(name_new) = sys.InputGroup.(name);
                sys.InputGroup = rmfield(sys.InputGroup, name);
                varargin{j} = sys;
                n_i = numel(sys.InputGroup.(name_new));
                sys_adds{j} = eye(n_i);
                ig_adds.(name_new) = idx_in + (1:n_i);
                idx_in = idx_in + n_i;
            end
        end

        sys_adds_ig{i} = ss(vertcat(sys_adds{:}));
        sys_adds_ig{i}.InputGroup.(name) = 1:n_i;
        sys_adds_ig{i}.OutputGroup = ig_adds;
        sys_adds_ig{i}.Ts = varargin{1}.Ts;
    end
end
for i = 1:numel(og_unique)
    og_adds = struct();
    is_overwrap = false(n, 1);
    for j = 1:n
        if sum(strcmp(og{j}, og_unique{i})) ~= 0
            is_overwrap(j) = true;
        end
    end
    if sum(is_overwrap) ~= 1
        n_i = nan;
        sys_adds = cell(n, 1);
        idx_out = 0;
        for j = 1:n
            if is_overwrap(j)
                name = og_unique{i};
                sys = varargin{j};
                name_new = strcat(name, '_output_', num2str(j));
                sys.OutputGroup.(name_new) = sys.OutputGroup.(name);
                sys.OutputGroup = rmfield(sys.OutputGroup, name);
                varargin{j} = sys;
                n_o = numel(sys.OutputGroup.(name_new));
                sys_adds{j} = eye(n_o);
                og_adds.(name_new) = idx_out + (1:n_o);
                idx_out = idx_out + n_o;
            end
        end

        sys_adds_og{i} = ss(horzcat(sys_adds{:}));
        sys_adds_og{i}.OutputGroup.(name) = 1:n_o;
        sys_adds_og{i}.InputGroup = og_adds;
        sys_adds_og{i}.Ts = varargin{1}.Ts;
    end
end

sys_all = blkdiag(varargin{1:end-1}, sys_adds_og{:}, sys_adds_ig{:});
ig = fieldnames(sys_all.InputGroup);
og = fieldnames(sys_all.OutputGroup);
idx_connect = intersect(ig, og);
if ~isempty(varargin{end})
    idx_connect = intersect(idx_connect, varargin{end});
end

for i = 1:numel(sys_adds_og)
    if ~isempty(sys_adds_og{i})
        idx_connect = union(idx_connect, fieldnames(sys_adds_og{i}.InputGroup));    
    end
end
for i = 1:numel(sys_adds_ig)
    if ~isempty(sys_adds_ig{i})
        idx_connect = union(idx_connect, fieldnames(sys_adds_ig{i}.OutputGroup));    
    end
end
% idx_connect = union(idx_connect, fieldnames(ig_adds));

feedin = tools.hcellfun(@(name) sys_all.InputGroup.(name), idx_connect);
feedout = tools.hcellfun(@(name) sys_all.OutputGroup.(name), idx_connect);

sys2 = ss(eye(numel(feedin)));
sys2.Ts = sys_all.Ts;

sys = feedback(sys_all, sys2, feedin, feedout, 1);
sys = sstools.fix_ss(sys);

end

