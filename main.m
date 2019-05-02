function main()
%Auther��Beth
%Time��2019-5-3
%Function:To draw a pixel painting with an Excel file by filling grids'
%background color.

clear
%Select an image
[file,path] = uigetfile('*.*','ѡ�����ͼƬ�����ļ�');
if isequal(file,0)
   disp('User selected Cancel');
   return
else
   disp(['User selected ', fullfile(path,file)]);
end
%Read an image
I=imread([path,file]);
[h0,w0,~]=size(I);
%I=imread('timg.jpg');
%�ɻ�������ͼƬ·��+����
%Optional.The image pathname+filename can be replaced.

msg1=sprintf('ԭͼƬ�߶�Ϊ%d,����ת��֮���ͼƬ�߶ȣ���λ�����أ�:',h0);
msg2=sprintf('ԭͼƬ���Ϊ%d,����ת��֮���ͼƬ��ȣ���λ�����أ�:',w0);
prompt = {msg1,msg2};
title = 'ͼƬ�����ش�СԽС�����ػ�Խ����';
dims = [1 35];

de1=sprintf('%d',h0/2);
de2=sprintf('%d',w0/2);
definput = {de1,de2};
clear de1,clear de2;
answer = inputdlg(prompt,title,dims,definput);
I=imresize(I,[str2double(answer{1}),str2double(answer{2})]);
%ͼƬ��С�任����ѡ��
%Optional.To resize the image.

%���ͼƬ�ǻҶ�ͼ����չ����ͨ��
%If the image is a grayscale image, expand it to three channels
if size(I,3)==1
    II(:,:,1)=I;
    II(:,:,2)=I;
    II(:,:,3)=I;
    I=II;
end

I=uint8(I);
[h,w,~]=size(I);

allCol=AllCol(w);
%�γ�w�е������б꣬�����Ԫ������
%Form all the column labels of all w columns, the output variable allCol is a CELL array

hExcel = actxserver('excel.application');   
% ����һ��excelʵ������
% Creat an excel COM
hWorkbooks = hExcel.Workbooks;     % ����һ��������������
hWorkbook = hWorkbooks.invoke('Add');    % ����һ��������(��)����
hSheets = hExcel.ActiveWorkBook.Sheets;      % ��õ�ǰ���������


handle = get(hSheets,'item',1);% Obtain the present sheet's handle
handle.Activate;
handle.Rows.RowHeight = 4;
handle.Columns.ColumnWidth =0.5;
set(hExcel,'Visible',1);

for ii=1:h
    for jj=1:w
        R=I(ii,jj,1);G=I(ii,jj,2);B=I(ii,jj,3);
        r=dec2hex(R);g=dec2hex(G);b=dec2hex(B);
        if length(r)==1
            r=['0',r];
        end
        if length(g)==1
            g=['0',g];
        end
        if length(b)==1
            b=['0',b];
        end
        range1=[allCol{jj},num2str(ii)];
        
        handle.Range(range1).Interior.Color=hex2dec([b,g,r]);
    end
end

hWorkbook.SaveAs('filename.xls');

delete(hExcel);

function allCol=AllCol(n)
%��������:�õ�1��n�е��б�
%n������������������
%allCol�������б꣬Ԫ�����飬ÿ��Ԫ����

allCol=cell(1,n);
Alpha='ABCDEFGHIJKLMNOPQRSTUVWXYZ';
col=1;
for N=1:n+round(n/10)
    k(1)=mod(N,27);
    ii=1;

    j(1)=N-k(1)*27^0;

    while j(ii)~=0
        k(ii+1)=mod(j(ii)/27^ii,27);
        j(ii+1)=j(ii)-k(ii+1)*27^ii;
        ii=ii+1;
    end
    
    if ~ismember(0,k)
        temp='';
        for ii=1:length(k)
           temp=[Alpha(k(ii)),temp];
        end
        
        allCol{col}=temp;
        col=col+1;
    end
        
end

