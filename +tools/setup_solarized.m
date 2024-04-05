function setup_solarized(type)
arguments
    type (1,1) string {mustBeMember(type,["light","dark", "default"])} = "dark"
end

colors_hex = struct();
colors_hex.base03 = "#002b36";
colors_hex.base02= "#073642";
colors_hex.base01= "#586e75";
colors_hex.base00= "#657b83";
colors_hex.base0= "#839496";
colors_hex.base1= "#93a1a1";
colors_hex.base2= "#eee8d5";
colors_hex.base3= "#fdf6e3";
colors_hex.yellow= "#b58900";
colors_hex.orange= "#cb4b16";
colors_hex.red= "#dc322f";
colors_hex.magenta= "#d33682";
colors_hex.violet= "#6c71c4";
colors_hex.blue= "#268bd2";
colors_hex.cyan= "#2aa198";
colors_hex.green = "#859900";

colors = tools.structfun(@(val) tools.hex2rgb(val, normalize=false), colors_hex);

s = settings;

s.matlab.colors.SystemCommandColor.PersonalValue = colors.base1;
s.matlab.colors.StringColor.PersonalValue = colors.base0;
end