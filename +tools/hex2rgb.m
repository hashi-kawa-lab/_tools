function rgb = hex2rgb(hex, options)
arguments
    hex
    options.normalize = true
end

    if startsWith(hex, '#')
        hex = replace(hex, "#", "");
    end
    rgb = sscanf(hex,'%2x%2x%2x',[1 3]);
    if options.normalize
        rgb = rgb / 255;
    end

end