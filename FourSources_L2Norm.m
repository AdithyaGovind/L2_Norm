clear vars
clc

Year=[2002,2003,2004,2005,2011,2012,2017,2019]; %Sample Years

cell=["A1","I1","Q1","Y1","AG1","AO1","AW1","BE1"]; %Excel Cells


for z=1:8
    
% Change year for different sample years 
    year=Year(z);

    %SOURCE FILE
    filename=sprintf('Source_%d',year);
    Source=xlsread('Dioxin_Master_WithoutHeaders.xlsx',filename);
    
    %%% Source file arrangement %%%
c=0;
AA=size(Source,1);
S=zeros(17,4);
% Arranging/forming the source matrix
for i1=1:17:AA
  c=c+1;
 
  for i2=i1:(i1+17)-1
    
      % S matrix: 17 congener rows, 4 sources=> columns
      
      S((i2-i1)+1,c)=Source(i2,1);
  end
end
 Source_row_size=size(S,1); 
 Source_col_size=size(S,2);

 S=S./sum(S,1);% Normalization 
    
    %SAMPLE CONCENTRATION FILE
    filename1=sprintf('SamplingConc_%d',year);
    Conc=xlsread('Dioxin_Master_WithoutHeaders.xlsx',filename1);
    
    
    %%% Concentration file arrangement %%%
c=0;
BB=size(Conc,1);
Sites=BB/17;
C=zeros(17,Sites);
%Reading concentration file
for i1=1:17:BB
    c=c+1;
    T=0;
    for i2=i1:(i1+17)-1
      T=T+1;
        % C matrix: 17 congener rows, (no. of col=no. of stns)
       
        C((i2-i1)+1,c)=Conc(i2,2);
    end 
end

 Conc_row_size=size(C,1); 
 Conc_col_size=size(C,2); 
 
 C=C./sum(C,1); % Normalization  



lb=[0;0;0;0];%Lower bound

Aeq=[1 1 1 1];% Equality constraint
beq=1;% Equality constraint


a=zeros(4,1);% Temporary array for alpha
alpha=zeros(Conc_col_size,4);


for j=1:Conc_col_size
    Con=zeros(1,17);% Temporary array to hold the conc. for a site
    for i=1:17
        Con(i)=C(i,j);
    end
    Concentration=Con'; % Define Concentration
    clear Con;
    
    %Optimization: Linear least squares solution
    [a,resnorm,residual,exitflag,output,lambda]= lsqlin(S,Concentration,[],[],Aeq,beq,lb);
    
    %Arranging the percent contributions for each site
    for k=1:4
    alpha(j,k)=a(k);
    end
    
end

% Output File
sheetname=sprintf('%d',year);
xlswrite('FourSources_L2_Norm.xlsx',alpha,sheetname);
end