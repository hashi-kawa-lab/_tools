classdef plot_pptx < handle

    properties
        h
        pages
        blankSlide
        myPres
        slide
        width = 960
        height = 460
        gap = 30
        header = 45
    end

    methods
        function obj = plot_pptx(name)
            if nargin < 1
                name = 'my_template_wide.pptx';
            end
            obj.h = actxserver('PowerPoint.Application');
            if exist(which(name), 'file')
                obj.myPres=obj.h.Presentations.Open(which(name));
                obj.pages = 1;
            else
                obj.myPres=obj.h.Presentations.Add;
                obj.pages = 0;
            end
            obj.blankSlide = obj.myPres.SlideMaster.CustomLayouts.Item(2);
            obj.add_page();
        end

        function add_page(obj, name)
            obj.pages = obj.pages + 1;
            obj.slide = obj.myPres.Slides.AddSlide(obj.pages, obj.blankSlide);
            if nargin > 1
                obj.slide.Shapes.Title.TextFrame.TextRange.Text = name;
            end
        end

        function add_plot(obj, f, row, col, row_all, col_all, retry, isequal)
            if nargin < 7 || isempty(retry)
                retry = 0;
            end
            if nargin < 8
                isequal = false;
            end
            if retry > 4
                return
            end
            try
                % width = [0 960], height = [45, 540-35]

                width = (obj.width-30*(col_all+1))/col_all;
                height = (460-30*(row_all+1))/row_all;

                figure(f), grid on, box on
                try
                    set(f.Children, 'FontSize', 10);
                    set(f.Children, 'FontName', 'Times New Roman');
                catch
                end
                set(f, 'Position', [100, 100, width, height])
                if isequal
                    axis('equal')
                end
                drawnow
                pause(1)
                copygraphics(f, 'ContentType', 'vector', 'BackgroundColor', 'none')
                Image = obj.slide.Shapes.Paste;
                set(Image, 'Left', 30*col+width*(col-1));
                set(Image, 'Top', 45+30*row+height*(row-1));
                set(Image, 'Width', width);
            catch
                obj.add_plot(f, row, col, row_all, col_all, retry+1)
            end
        end

        function add_plot2(obj, f, row, col, row_all, col_all, retry, isequal)
            if nargin < 7 || isempty(retry)
                retry = 0;
            end
            if nargin < 8
                isequal = false;
            end
            if retry > 4
                return
            end
            try
                % width = [0 960], height = [45, 540-35]

                width = (obj.width-30*(col_all+1))/col_all;
                height = (460-30*(row_all+1))/row_all;

                figure(f), grid on, box on
                try
                    set(f.Children, 'FontSize', 10);
                    set(f.Children, 'FontName', 'Times New Roman');
                catch
                end
                set(f, 'Position', [100, 100, width, height])
                if isequal
                    axis('equal')
                end
                copygraphics(f, 'ContentType', 'image', 'BackgroundColor', 'none')
                Image = obj.slide.Shapes.Paste;
                set(Image, 'Left', 30*col+width*(col-1));
                set(Image, 'Top', 45+30*row+height*(row-1));
                set(Image, 'Width', width);
            catch
                obj.add_plot(f, row, col, row_all, col_all, retry+1)
            end
        end

        function add_plot3(obj, f, row, col, row_all, col_all, retry, isequal)
            if nargin < 7 || isempty(retry)
                retry = 0;
            end
            if nargin < 8
                isequal = false;
            end
            if retry > 4
                return
            end
            try
                width = (obj.width - 30 * (col_all + 1)) / col_all;
                height = (460 - 30 * (row_all + 1)) / row_all;

                figure(f), grid on, box on
                try
                    set(f.Children, 'FontSize', 10);
                    set(f.Children, 'FontName', 'Times New Roman');
                catch
                end
                set(f, 'Position', [100, 100, width, height])
                if isequal
                    axis('equal')
                end
                drawnow
                pause(1)

                % Step 1: ベクターで一時的に貼り付け（文字化け回避）
                copygraphics(f, 'ContentType', 'vector', 'BackgroundColor', 'none');
                tempShape = obj.slide.Shapes.Paste;
                tempShape.Left = 30 * col + width * (col - 1);
                tempShape.Top = 45 + 30 * row + height * (row - 1);
                tempShape.Width = width;

                % Step 2: PowerPoint内でコピー → 図（PNG）として貼り付け
                tempShape.Copy;
                pause(0.5);  % 貼り付け前に待機
                finalShape = obj.slide.Shapes.PasteSpecial('ppPastePNG');
                finalShape.Left = tempShape.Left;
                finalShape.Top = tempShape.Top;
                finalShape.Width = width;

                % keyboard
                % Step 3: 中間のベクター図形を削除
                tempShape.Delete;

            catch
                obj.add_plot3(f, row, col, row_all, col_all, retry + 1, isequal)
            end
        end

        function add_plot_position(obj, f, left, top, width, height, options, NamedArgs)
            arguments
                obj
                f
                left = []
                top = []
                width = []
                height = []
                options.type (1,:) char {mustBeMember(options.type, {'vector', 'image'})} = 'vector'
                options.fontsize = 18;
                options.fontname = "Times New Roman"
                options.resolution (1,1) double {mustBePositive, mustBeInteger} = 300
                NamedArgs.left
                NamedArgs.top
                NamedArgs.width
                NamedArgs.height
            end

            figure(f), grid on, box on
            set(f.Children, 'FontSize', options.fontsize);
            set(f.Children, 'FontName', options.fontname);
            if isempty(left)
                left = NamedArgs.left;
            end
            if isempty(top)
                top = NamedArgs.top;
            end
            if isempty(width)
                width = NamedArgs.width;
            end
            if isempty(height)
                height = NamedArgs.height;
            end
            tmpfile = tempname(tempdir);
            if strcmpi(options.type, 'vector')
                tmpfile = strcat(tmpfile, '.emf');
                try
                    exportgraphics(f, tmpfile, ContentType="vector", BackgroundColor="none", Width=width, Height=height);
                catch
                    set(f, 'Position', [100, 100, width, height])
                    exportgraphics(f, tmpfile, ContentType="vector", BackgroundColor="none");
                end
            else
                tmpfile = strcat(tmpfile, '.png');
                try
                    exportgraphics(f, tmpfile, ContentType="image", BackgroundColor="none", Width=width, Height=height, Units="points", Resolution=options.resolution);
                catch
                    set(f, 'Position', [100, 100, width, height])
                    exportgraphics(f, tmpfile, ContentType="image", BackgroundColor="none", Resolution=options.resolution);
                end
            end

            obj.slide.Shapes.AddPicture(tmpfile, 'msoFalse', 'msoCTrue', left, top, width, height);
            delete(tmpfile);

        end

        function add_plot_tile(obj, f, row, col, row_all, col_all, options, common_options)
            arguments
                obj
                f
                row (1,1) {mustBePositive, mustBeInteger}
                col (1,1) {mustBePositive, mustBeInteger}
                row_all (1,1) {mustBePositive, mustBeInteger}
                col_all (1,1) {mustBePositive, mustBeInteger}
                options.gap = obj.gap
                options.header = obj.header
                options.slidewidth = obj.width
                options.slideheight = obj.height
                common_options.type (1,:) char {mustBeMember(common_options.type, {'vector', 'image'})} = 'vector'
                common_options.fontsize = 18;
                common_options.fontname = "Times New Roman"
                common_options.resolution (1,1) double {mustBePositive, mustBeInteger} = 300
            end

            width = (options.slidewidth - options.gap * (col_all + 1)) / col_all;
            height = (options.slideheight - options.gap * (row_all + 1)) / row_all;
            left = options.gap * col + width * (col - 1);
            top = options.header + options.gap * row + height * (row - 1);
            common_options = tools.struct2parameter(common_options);
            obj.add_plot_position(f, left, top, width, height, common_options{:});

        end

        function align_figs(obj)
            rows = 1;
            columns = 1;
            % すべてのfigureのハンドルを取得
            allFigures = findall(0, 'Type', 'figure');
            numbers = {allFigures.Number};
            idx = tools.vcellfun(@(c) ~isempty(c), numbers);
            allFigures = allFigures(idx);
            numbers = vertcat(numbers{idx});
            [~, i] = sort(numbers);
            allFigures = allFigures(i);
            % figureの数
            numFigures = numel(allFigures);
            % 行数と列数の積がfigureの数より大きい場合、行数と列数を調整
            while rows * columns < numFigures
                if rows < columns
                    rows = rows + 1;
                else
                    columns = columns + 1;
                end
            end

            k = 1;
            for i = 1:rows
                for j = 1:columns
                    fig = allFigures(k);
                    obj.add_plot_tile(fig, i, j, rows, columns);
                    k = k + 1;
                    if k > numFigures
                        tools.figs2front;
                        return;
                    end
                end
            end

        end


        function add_text(obj, text, left, top, width, height, options, NamedArg)
            arguments
                obj
                text
                left = []
                top = []
                width = []
                height = []
                options.bold = false
                options.italic = false
                options.underline = false
                options.shadow = false
                options.fontname
                options.fontsize
                options.autosize = true
                options.wordwrap = false
                NamedArg.left
                NamedArg.top
                NamedArg.width = 1
                NamedArg.height = 1
            end
            if isempty(left)
                left = NamedArg.left;
            end
            if isempty(top)
                top = NamedArg.top;
            end
            if isempty(width)
                width = NamedArg.width;
            end
            if isempty(height)
                height = NamedArg.height;
            end
            textbox = obj.slide.Shapes.AddTextbox('msoTextOrientationHorizontal', left, top, width, height);
            if options.autosize
                textbox.TextFrame.AutoSize = 'ppAutoSizeShapeToFitText';
            else
                textbox.TextFrame.AutoSize = 'ppAutoSizeNone';
            end
            if options.wordwrap
                textbox.TextFrame.WordWrap = 'msoTrue';
            else
                textbox.TextFrame.WordWrap = 'msoFalse';
            end
            textbox.TextFrame.TextRange.Text = text;
            textbox.TextFrame.TextRange.Font.Bold = options.bold;
            textbox.TextFrame.TextRange.Font.Italic = options.italic;
            textbox.TextFrame.TextRange.Font.Underline = options.underline;
            textbox.TextFrame.TextRange.Font.Shadow = options.shadow;
            if isfield(options, 'fontsize')
                textbox.TextFrame.TextRange.Font.Size = options.fontsize;
            end
            if isfield(options, 'fontname')
                textbox.TextFrame.TextRange.Font.Name = options.fontname;
            end
        end

        function close(obj)
            Quit(obj.h)
            delete(obj.h)
        end

        function saveas(obj, name)
            obj.myPres.SaveAs(name);
        end
    end

end