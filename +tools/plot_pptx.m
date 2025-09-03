classdef plot_pptx < handle
    
    properties
        h
        pages
        blankSlide
        myPres
        slide
        width = 960
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

        
        function close(obj)
            Quit(obj.h)
            delete(obj.h)
        end
        
        function saveas(obj, name)
           obj.myPres.SaveAs(name); 
        end
    end
    
end