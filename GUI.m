function varargout = GUI(varargin)
% GUI MATLAB code for GUI.fig
%      GUI, by itself, creates a new GUI or raises the existing
%      singleton*.
%
%      H = GUI returns the handle to a new GUI or the handle to
%      the existing singleton*.
%
%      GUI('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in GUI.M with the given input arguments.
%
%      GUI('Property', 'Value',...) creates a new GUI or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before GUI_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to GUI_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help GUI

% Last Modified by GUIDE v2.5 28-Dec-2024 10:54:13

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
    'gui_Singleton',  gui_Singleton, ...
    'gui_OpeningFcn', @GUI_OpeningFcn, ...
    'gui_OutputFcn',  @GUI_OutputFcn, ...
    'gui_LayoutFcn',  [] , ...
    'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before GUI is made visible.
function GUI_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to GUI (see VARARGIN)

% Choose default command line output for GUI
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);
axes(handles.axes7);
imshow('Cairo_University_crest.png');
axes(handles.axes8);
imshow('CUFE.jpg');

% UIWAIT makes GUI wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = GUI_OutputFcn(hObject, eventdata, handles)
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;



function fileNamePath_Callback(hObject, eventdata, handles)
% hObject    handle to fileNamePath (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of fileNamePath as text
%        str2double(get(hObject,'String')) returns contents of fileNamePath as a double


% --- Executes during object creation, after setting all properties.
function fileNamePath_CreateFcn(hObject, eventdata, handles)
% hObject    handle to fileNamePath (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end


% --- Executes on button press in browseButton.
function browseButton_Callback(hObject, eventdata, handles)
% hObject    handle to browseButton (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
[filename pathname] = uigetfile({'*.xlsx'},'File Selector'); % Get File Path and Name
fullpathname = strcat(pathname,filename); % Concatenate File Path and Name
datasheet = readcell(fullpathname);% Read File
datasheet = datasheet(2:end,:);
datasheet = sortrows(datasheet, 3);
assignin('base', 'datasheet',datasheet); % Assign Global Variable
set(handles.fileNamePath, 'String', fullpathname);



function voltage_Callback(hObject, eventdata, handles)
% hObject    handle to voltage (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of voltage as text
%        str2double(get(hObject,'String')) returns contents of voltage as a double


% --- Executes during object creation, after setting all properties.
function voltage_CreateFcn(hObject, eventdata, handles)
% hObject    handle to voltage (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end



function frequency_Callback(hObject, eventdata, handles)
% hObject    handle to frequency (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of frequency as text
%        str2double(get(hObject,'String')) returns contents of frequency as a double


% --- Executes during object creation, after setting all properties.
function frequency_CreateFcn(hObject, eventdata, handles)
% hObject    handle to frequency (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end



function bundles_Callback(hObject, eventdata, handles)
% hObject    handle to bundles (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of bundles as text
%        str2double(get(hObject,'String')) returns contents of bundles as a double


% --- Executes during object creation, after setting all properties.
function bundles_CreateFcn(hObject, eventdata, handles)
% hObject    handle to bundles (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end



function power_Callback(hObject, eventdata, handles)
power_Variable = get(handles.power, 'String');
if ((isempty(str2num(power_Variable))) || (str2num(power_Variable) < 0))
    set(handles.Power_Error_Text, 'String', 'Please write a valid value')
    assignin('base', 'Error_State', 1)
else
    set(handles.Power_Error_Text, 'String', '')
    assignin('base', 'Error_State', 0)
    assignin('base', 'Power', str2num(power_Variable));
end
% hObject    handle to power (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of power as text
%        str2double(get(hObject,'String')) returns contents of power as a double


% --- Executes during object creation, after setting all properties.
function power_CreateFcn(hObject, eventdata, handles)
% hObject    handle to power (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end



function edit6_Callback(hObject, eventdata, handles)
% hObject    handle to edit6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit6 as text
%        str2double(get(hObject,'String')) returns contents of edit6 as a double


% --- Executes during object creation, after setting all properties.
function edit6_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end



function lineLength_Callback(hObject, eventdata, handles)
length_Variable = get(handles.lineLength, 'String');
if ((isempty(str2num(length_Variable))) || (str2num(length_Variable) < 0))
    set(handles.Length_Error_Text, 'String', 'Please write a valid value')
    assignin('base', 'Error_State', 1)
else
    set(handles.Length_Error_Text, 'String', '')
    assignin('base', 'Error_State', 0)
    assignin('base', 'Line_Length', str2num(length_Variable))
end
% hObject    handle to lineLength (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
lineLength = str2num(get(handles.lineLength, 'String'));
if lineLength >= 80
    set(handles.piModelRadioButton, 'Enable', 'On', 'Value', 0);
    set(handles.tModelRadioButton, 'Enable', 'On', 'Value', 0);
else
    set(handles.piModelRadioButton, 'Enable', 'Off', 'Value', 0);
    set(handles.tModelRadioButton, 'Enable', 'Off', 'Value', 0);
end
% Hints: get(hObject,'String') returns contents of lineLength as text
%        str2double(get(hObject,'String')) returns contents of lineLength as a double


% --- Executes during object creation, after setting all properties.
function lineLength_CreateFcn(hObject, eventdata, handles)
% hObject    handle to lineLength (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end



function bundlesDistance_Callback(hObject, eventdata, handles)
bundleLength_Variable = get(handles.bundlesDistance, 'String');
if ((isempty(str2num(bundleLength_Variable))) || (str2num(bundleLength_Variable) < 0))
    set(handles.Bundles_Distance_Error_Text, 'String', 'Please write a valid value')
    assignin('base', 'Error_State', 1)
else
    set(handles.Bundles_Distance_Error_Text, 'String', '')
    assignin('base', 'Error_State', 0)
    assignin('base', 'Bundle_Distance', str2num(bundleLength_Variable))
end
% hObject    handle to bundlesDistance (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of bundlesDistance as text
%        str2double(get(hObject,'String')) returns contents of bundlesDistance as a double


% --- Executes during object creation, after setting all properties.
function bundlesDistance_CreateFcn(hObject, eventdata, handles)
% hObject    handle to bundlesDistance (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end



function cableDistance_Callback(hObject, eventdata, handles)
Distance_Variable = get(handles.cableDistance, 'String');
if ((isempty(str2num(Distance_Variable))) || (str2num(Distance_Variable) < 0))
    set(handles.Distance_Error_Text, 'String', 'Please write a valid value')
    assignin('base', 'Error_State', 1)
else
    set(handles.Distance_Error_Text, 'String', '')
    assignin('base', 'Error_State', 0)
    set(handles.equilateralButton, 'Enable', 'On', 'Value', 0)
    set(handles.horizontalButton, 'Enable', 'On', 'Value', 0)
    assignin('base', 'Distance', str2num(Distance_Variable))
end
% hObject    handle to cableDistance (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% Hints: get(hObject,'String') returns contents of cableDistance as text
%        str2double(get(hObject,'String')) returns contents of cableDistance as a double


% --- Executes during object creation, after setting all properties.
function cableDistance_CreateFcn(hObject, eventdata, handles)
% hObject    handle to cableDistance (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end



function bcDistance_Callback(hObject, eventdata, handles)
% hObject    handle to bcDistance (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of bcDistance as text
%        str2double(get(hObject,'String')) returns contents of bcDistance as a double


% --- Executes during object creation, after setting all properties.
function bcDistance_CreateFcn(hObject, eventdata, handles)
% hObject    handle to bcDistance (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end



function edit12_Callback(hObject, eventdata, handles)
% hObject    handle to edit12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit12 as text
%        str2double(get(hObject,'String')) returns contents of edit12 as a double


% --- Executes during object creation, after setting all properties.
function edit12_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end


% --- Executes on selection change in voltageMenu.
function voltageMenu_Callback(hObject, eventdata, handles)
% hObject    handle to voltageMenu (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
voltageValue = get(handles.voltageMenu, 'Value');
switch voltageValue
    case 1
        set(handles.twoBundles,'Enable', 'Off', 'Value', 0);
        set(handles.threeBundles,'Enable', 'Off', 'Value', 0);
        set(handles.fourBundles,'Enable', 'Off', 'Value', 0);
        assignin('base', 'voltageImpedance', 0);
        assignin('base', 'Voltage_Level', 0);
        set(handles.radiobutton9, 'Enable', 'Off', 'Value', 0);
        set(handles.radiobutton10, 'Enable', 'Off', 'Value', 0);
    case 2
        set(handles.twoBundles,'Enable', 'On', 'Value', 0);
        set(handles.threeBundles,'Enable', 'Off', 'Value', 0);
        set(handles.fourBundles,'Enable', 'Off', 'Value', 0);
        assignin('base', 'voltageImpedance', 0.3);
        assignin('base', 'Voltage_Level', 345);
        set(handles.radiobutton9, 'Enable', 'On', 'Value', 0);
        set(handles.radiobutton10, 'Enable', 'On', 'Value', 0);
    case 3
        set(handles.twoBundles,'Enable', 'On', 'Value', 0);
        set(handles.threeBundles,'Enable', 'On', 'Value', 0);
        set(handles.fourBundles,'Enable', 'On', 'Value', 0);
        assignin('base', 'voltageImpedance', 0.32);
        assignin('base', 'Voltage_Level', 500);
        set(handles.radiobutton9, 'Enable', 'On', 'Value', 0);
        set(handles.radiobutton10, 'Enable', 'On', 'Value', 0);
    case 4
        set(handles.twoBundles,'Enable', 'Off', 'Value', 0);
        set(handles.threeBundles,'Enable', 'On', 'Value', 0);
        set(handles.fourBundles,'Enable', 'On', 'Value', 0);
        assignin('base', 'voltageImpedance', 0.27);
        assignin('base', 'Voltage_Level', 765);
        set(handles.radiobutton9, 'Enable', 'On', 'Value', 0);
        set(handles.radiobutton10, 'Enable', 'On', 'Value', 0);
end
% Hints: contents = cellstr(get(hObject,'String')) returns voltageMenu contents as cell array
%        contents{get(hObject,'Value')} returns selected item from voltageMenu


% --- Executes during object creation, after setting all properties.
function voltageMenu_CreateFcn(hObject, eventdata, handles)
% hObject    handle to voltageMenu (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end


% --- Executes on button press in horizontalButton.
function horizontalButton_Callback(hObject, eventdata, handles)
% hObject    handle to horizontalButton (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
horizontalButtonValue = get(handles.horizontalButton, 'Value');
if horizontalButtonValue == 1
    assignin('base', 'abDistance', str2num(get(handles.cableDistance, 'String')));
    assignin('base', 'bcDistance', str2num(get(handles.cableDistance, 'String')));
    assignin('base', 'caDistance', 2*str2num(get(handles.cableDistance, 'String')));
end

% Hint: get(hObject,'Value') returns toggle state of horizontalButton


% --- Executes on button press in equilateralButton.
function equilateralButton_Callback(hObject, eventdata, handles)
% hObject    handle to equilateralButton (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
equilateralButtonValue = get(handles.equilateralButton, 'Value');
if equilateralButtonValue == 1
    assignin('base', 'abDistance', str2num(get(handles.cableDistance, 'String')));
    assignin('base', 'bcDistance', str2num(get(handles.cableDistance, 'String')));
    assignin('base', 'caDistance', str2num(get(handles.cableDistance, 'String')));
end
% Hint: get(hObject,'Value') returns toggle state of equilateralButton


% --- Executes during object creation, after setting all properties.
function cableName_CreateFcn(hObject, eventdata, handles)
% hObject    handle to cableName (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called


% --- Executes on button press in calculateButton.
function calculateButton_Callback(hObject, eventdata, handles)
Error_State = evalin('base', 'Error_State');
if ~(Error_State)
    %Retrieving Variables
    Power_Receiving = evalin('base', 'Power');
    Voltage_Level = evalin('base', 'Voltage_Level');
    PF = evalin('base', 'PF');
    Freq = evalin('base', 'Frequency');
    Voltage_Impedance = evalin('base', 'voltageImpedance');
    Model = evalin('base', 'Model');
    Line_Length = evalin('base', 'Line_Length');
    abDistance = evalin('base', 'abDistance');
    bcDistance = evalin('base', 'bcDistance');
    caDistance = evalin('base', 'caDistance');
    Bundle_Distance = evalin('base', 'Bundle_Distance');
    Bundles_Number = evalin('base', 'Bundles_Number');
    required_efficiency = evalin('base', 'Required_Efficiency');
    required_volt_regulation = evalin('base', 'Required_VR');
    
    %Calculations
    Estimate_Line_Current = Power_Receiving*1e6./Voltage_Level/1e3/sqrt(3)/PF;
    Circuits_Number = Power_Receiving*1e6*Line_Length*Voltage_Impedance/sin(30*pi/180)./(Voltage_Level*1e3)./(Voltage_Level*1e3);
    Circuits_Number = ceil(Circuits_Number);
    Circuits_Number = Circuits_Number+1;
    Estimate_CCC = Estimate_Line_Current/(Circuits_Number * Bundles_Number);
    
    datasheet = evalin('base', 'datasheet');
    
    Cable_Names = cellstr(datasheet(:,1));
    assignin('base', 'Cable_Names', Cable_Names);
    Cable_GMR = cell2mat(datasheet(:,2));
    assignin('base', 'Cable_GMR', Cable_GMR);
    Cable_CCC = cell2mat(datasheet(:,3));
    assignin('base', 'Cable_CCC', Cable_CCC);
    Cable_Diameter = cell2mat(datasheet(:,4));
    assignin('base', 'Cable_Diameter', Cable_Diameter);
    Cable_Rac60 = cell2mat(datasheet(:,5));
    assignin('base', 'Cable_Rac60', Cable_Rac60);
    Cable_GMR = Cable_GMR./1000;
    Cable_Radius = Cable_Diameter./2000;
    
    D_equivalent = (abDistance * bcDistance * caDistance)^(1/3);
    D_equivalent_Array = D_equivalent.*linspace(0.5, 2, 1000);
    
    %Finding Starting Index
    Cable_index = 1;
    Loop_Flag = 1;
    while Loop_Flag
        if (Estimate_CCC < Cable_CCC(Cable_index))
            Loop_Flag = 0;
            %Cable_Starting_Index = Cable_index;
        elseif (Cable_index == length(Cable_CCC))
            Loop_Flag = 0;
            %Cable_Starting_Index=length(Cable_CCC);
        end
        Cable_index = Cable_index + 1;
    end
    Cable_Starting_Index = Cable_index - 1;
    
    Found = 0;
    %Find Required Efficiency and Voltage Regulation
    while ~Found
        %Calculate GMRs
        switch Bundles_Number
            case 2
                GMR_L = (Cable_GMR(Cable_Starting_Index).*Bundle_Distance).^0.5;
                GMR_C = (Cable_Radius(Cable_Starting_Index).*Bundle_Distance).^0.5;
            case 3
                GMR_L = (Cable_GMR(Cable_Starting_Index).*Bundle_Distance^2).^(1/3);
                GMR_C = (Cable_Radius(Cable_Starting_Index).*Bundle_Distance^2).^(1/3);
            case 4
                GMR_L = (Cable_GMR(Cable_Starting_Index).*sqrt(2)*Bundle_Distance^3).^(1/4);
                GMR_C = (Cable_Radius(Cable_Starting_Index).*sqrt(2)*Bundle_Distance^3).^(1/4);
        end
        %Calculate Circuit Parameters
        Xl = 2*pi*Freq*1e3*2*1e-7*(log(1/GMR_L)+log(D_equivalent)); %XL per phase in ohm/km
        E0 = 8.85*1e-12; % epslon not
        Xc = ((1/(2*pi*Freq*2*pi*E0))*(log(1/GMR_C)+log(D_equivalent))).*1e3*1e-6; %Xc per phase in Mohm/km
        R = Cable_Rac60(Cable_Starting_Index)*1e-3*Line_Length*sqrt(Freq/60);
        Xc = (Xc*1e6*Line_Length*(60/Freq))/Circuits_Number;
        Xl = Xl*Line_Length*(Freq/60)/Circuits_Number;
        R = R/(Circuits_Number*Bundles_Number);
        Z = R+1i*Xl;
        Y = (1i*1./Xc);
        Zo = sqrt(Z/Y);
        Yo = sqrt(Z*Y);
        Ir = Estimate_Line_Current*(PF-1i*sin(acos(PF)));
        if Line_Length<80 && Line_Length>0 %Short
            A = 1;
            B = Z;
            C = 0;
            D = 1;
            Current_Sending = C*Voltage_Level*1e3/sqrt(3)+D*Ir;
            Voltage_Sending = Voltage_Level*1e3*A/sqrt(3)+B*Ir;
            vro = abs(Voltage_Sending/A/1e3*sqrt(3));
        elseif Line_Length<240 && Line_Length>=80 && get(handles.tModelRadioButton, 'Value') %medium T model
            A = (1+Y.*Z/2);
            B = Z.*(1+Z.*Y/4);
            C = Y;
            D = A;
            Current_Sending = C*Voltage_Level*1e3/sqrt(3)+D*Ir;
            Voltage_Sending = Voltage_Level*1e3*A/sqrt(3)+B*Ir;
            vro = abs(Voltage_Sending/A/1e3*sqrt(3)); %no load line voltage in kv
        elseif Line_Length<240 && Line_Length>=80 && get(handles.piModelRadioButton, 'Value')
            %pi model
            A = (1+Y.*Z/2);
            B = Z;
            C = Y.*(1+Z.*Y/4);
            D = A;
            Current_Sending = C*Voltage_Level*1e3/sqrt(3)+D*Ir;
            Voltage_Sending = Voltage_Level*1e3*A/sqrt(3)+B*Ir;
            vro = abs(Voltage_Sending/A/1e3*sqrt(3));
        elseif Line_Length>= 240 && get(handles.tModelRadioButton, 'Value') % Long T model
            Y_dash = sinh(Yo)/Zo;
            Z_dash = (2*cosh(Yo)-2)/Y_dash;
            A = (1+(Y_dash*Z_dash)/2);
            B = Z_dash*(1+(Y_dash*Z_dash)/4);
            C = Y_dash;
            D = (1+(Y_dash*Z_dash)/2);
            Voltage_Sending = A*Voltage_Level*1e3/sqrt(3) + B*Ir;
            Current_Sending = C*Voltage_Level*1e3/sqrt(3) + D*Ir;
            vro = abs(Voltage_Sending/A/1e3*sqrt(3)); % no load
        elseif Line_Length>= 240 && get(handles.piModelRadioButton, 'Value') % Long Pi model
            Z_dash = Zo*sinh(Yo);
            Y_dash = (2*cosh(Yo)-2)/Z_dash;
            A = (1+(Z_dash*Y_dash)/2);
            B = Z_dash;
            C = Y_dash*(1+(Z_dash*Y_dash)/4);
            D = (1+(Z_dash*Y_dash)/2);
            Voltage_Sending = A*Voltage_Level*1e3/sqrt(3) + B*Ir;
            Current_Sending = C*Voltage_Level*1e3/sqrt(3) + D*Ir;
            vro = abs(Voltage_Sending/A/1e3*sqrt(3)); % no load
        end
        Calculated_VR = (vro-Voltage_Level)/(Voltage_Level)*100;
        Power_Sending = 3*real(Voltage_Sending*conj(Current_Sending))/1e6;
        Power_losses = Power_Sending-Power_Receiving;
        Calculated_Efficiency = (Power_Receiving/Power_Sending)*100;
        if (Calculated_VR<required_volt_regulation && Calculated_Efficiency>required_efficiency)
            Found = 1;
            material_number = Cable_Starting_Index;
            N = 0;
        elseif Cable_Starting_Index == (length(datasheet))
            Found = 1;
            material_number = Cable_Starting_Index;
            N = 1;
        end
        Cable_Starting_Index = Cable_Starting_Index + 1;
    end
    %%Output displayed for user
    if N==0
        set(handles.suggestedCableName, 'String', Cable_Names(material_number));
        set(handles.suggestedCableCCC, 'String', num2str(Cable_CCC(material_number)));
        set(handles.calculatedEfficiency, 'String', num2str(Calculated_Efficiency));
        set(handles.calculatedVR, 'String', num2str(Calculated_VR));
        set(handles.calculatedCCC_Static, 'String', num2str(Estimate_CCC));
        set(handles.calculatedLosses, 'String', num2str(Power_losses));
        set(handles.Min_Circuits_Number, 'String', num2str(Circuits_Number));
        set(handles.noCable, 'String', '');
        
    elseif N==1
        set(handles.noCable, 'String', 'No cable with the desired values was found. The values shown are for the last element in the datasheet sorted by Current Carrying Capacity.');
        set(handles.suggestedCableName, 'String', Cable_Names(material_number));
        set(handles.suggestedCableCCC, 'String', num2str(Cable_CCC(material_number)));
        set(handles.calculatedEfficiency, 'String', num2str(Calculated_Efficiency));
        set(handles.calculatedVR, 'String', num2str(Calculated_VR));
        set(handles.calculatedCCC_Static, 'String', num2str(Estimate_CCC));
        set(handles.calculatedLosses, 'String', num2str(Power_losses));
        set(handles.Min_Circuits_Number, 'String', num2str(Circuits_Number));
    end
    for i = 1:length(D_equivalent_Array)
        Xl_Array(i) = 2*pi*Freq*1e3*2*1e-7.*(log(1/GMR_L)+log(D_equivalent_Array(i))); %XL per phase in ohm/km
        Xc_Array(i) = ((1/(2*pi*Freq*2*pi*E0)).*(log(1/GMR_C)+log(D_equivalent_Array(i)))).*1e3*1e-6; %Xc per phase in Mohm/km
        R = Cable_Rac60(material_number)*1e-3*Line_Length*sqrt(Freq/60);
        Xc_Array(i) = (Xc_Array(i).*1e6*Line_Length*sqrt(Freq/60))/Circuits_Number;
        Xl_Array(i) = Xl_Array(i).*Line_Length*sqrt(Freq/60)/Circuits_Number;
        R = R/(Circuits_Number*Bundles_Number);
        Z_Array(i) = R+1i.*Xl_Array(i);
        Y_Array(i) = (1i.*1./Xc_Array(i));
        %Long Model Calculations
        Zo_Array(i) = sqrt(Z_Array(i)./Y_Array(i));
        Yo_Array(i) = sqrt(Z_Array(i).*Y_Array(i));
        Ir = Estimate_Line_Current*(PF-1i*sin(acos(PF)));
        if Line_Length<80 && Line_Length>0 %Short
            A_Array(i) = 1;
            B_Array(i) = Z_Array(i);
            C_Array(i) = 0;
            D_Array(i) = 1;
            Current_Sending_Array(i) = C_Array(i).*Voltage_Level*1e3/sqrt(3)+D_Array(i).*Ir;
            Voltage_Sending_Array(i) = Voltage_Level*1e3.*A_Array(i)./sqrt(3)+B_Array(i).*Ir;
            vro_Array(i) = abs(Voltage_Sending_Array(i)./A_Array(i)./1e3*sqrt(3));
        elseif Line_Length<240 && Line_Length>=80 && get(handles.tModelRadioButton, 'Value') %medium T model
            A_Array(i) = (1+Y_Array(i).*Z_Array(i)./2);
            B_Array(i) = Z_Array(i).*(1+Z_Array(i).*Y_Array(i)./4);
            C_Array(i) = Y_Array(i);
            D_Array(i) = A_Array(i);
            Current_Sending_Array(i) = C_Array(i).*Voltage_Level*1e3/sqrt(3)+D_Array(i).*Ir;
            Voltage_Sending_Array(i) = Voltage_Level*1e3.*A_Array(i)./sqrt(3)+B_Array(i).*Ir;
            vro_Array(i) = abs(Voltage_Sending_Array(i)./A_Array(i)./1e3*sqrt(3)); %no load line voltage in kv
        elseif Line_Length<240 && Line_Length>=80 && get(handles.piModelRadioButton, 'Value')
            %pi model
            A_Array(i) = (1+Y_Array(i).*Z_Array(i)/2);
            B_Array(i) = Z_Array(i);
            C_Array(i) = Y_Array(i).*(1+Z_Array(i).*Y_Array(i)/4);
            D_Array(i) = A_Array(i);
            Current_Sending_Array(i) = C_Array(i).*Voltage_Level*1e3/sqrt(3)+D_Array(i).*Ir;
            Voltage_Sending_Array(i) = Voltage_Level*1e3.*A_Array(i)./sqrt(3)+B_Array(i).*Ir;
            vro_Array(i) = abs(Voltage_Sending_Array(i)./A_Array(i)./1e3*sqrt(3));
        elseif Line_Length>= 240 && get(handles.tModelRadioButton, 'Value') %long & T model
            Y_dash_Array(i) = sinh(Yo_Array(i))./Zo_Array(i);
            Z_dash_Array(i) = (2.*cosh(Yo_Array(i))-2)./Y_dash_Array(i);
            A_Array(i) = (1+(Y_dash_Array(i).*Z_dash_Array(i))./2);
            B_Array(i) = Z_dash_Array(i).*(1+(Y_dash_Array(i).*Z_dash_Array(i))./4);
            C_Array(i) = Y_dash_Array(i);
            D_Array(i) = (1+(Y_dash_Array(i).*Z_dash_Array(i))./2);
            Voltage_Sending_Array(i) = A_Array(i).*Voltage_Level.*1e3./sqrt(3) + B_Array(i).*Ir;
            Current_Sending_Array(i) = C_Array(i).*Voltage_Level.*1e3./sqrt(3) + D_Array(i).*Ir;
            vro_Array(i) = abs(Voltage_Sending_Array(i)./A_Array(i)./1e3.*sqrt(3)); % no load
        elseif Line_Length>= 240 && get(handles.piModelRadioButton, 'Value') %long & pi model
            Z_dash_Array(i) = Zo_Array(i).*sinh(Yo_Array(i));
            Y_dash_Array(i) = (2.*cosh(Yo_Array(i))-2)./Z_dash_Array(i);
            A_Array(i) = (1+(Z_dash_Array(i).*Y_dash_Array(i))./2);
            B_Array(i) = Z_dash_Array(i);
            C_Array(i) = Y_dash_Array(i).*(1+(Z_dash_Array(i).*Y_dash_Array(i))./4);
            D_Array(i) = (1+(Z_dash_Array(i).*Y_dash_Array(i))./2);
            Voltage_Sending_Array(i) = A_Array(i).*Voltage_Level.*1e3./sqrt(3) + B_Array(i).*Ir;
            Current_Sending_Array(i) = C_Array(i).*Voltage_Level.*1e3./sqrt(3) + D_Array(i).*Ir;
            vro_Array(i) = abs(Voltage_Sending_Array(i)./A_Array(i)./1e3.*sqrt(3)); % no load
        end
        Calculated_VR_Array(i) = (vro_Array(i)-Voltage_Level)./(Voltage_Level).*100;
        Power_Sending_Array(i) = 3*real(Voltage_Sending_Array(i).*conj(Current_Sending_Array(i)))./1e6;
        Power_losses_Array(i) = Power_Sending_Array(i)-Power_Receiving;
        Calculated_Efficiency_Array(i) = (Power_Receiving./Power_Sending_Array(i)).*100;
    end
    axes(handles.axes12);
    plot(D_equivalent_Array, Calculated_Efficiency_Array);
    grid on;
    title('Effect of distance between cables on efficiency');
    xlabel('Distance between cables','color','r')
    ylabel('Efficency change','color','r')
    axes(handles.axes13);
    plot(D_equivalent_Array, Calculated_VR_Array);
    grid on;
    title('Effect of distance between cables on voltage regulation');
    xlabel('Distance between cables','color','r')
    ylabel('Voltage Regulation change','color','r')
end


% hObject    handle to calculateButton (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)




function edit14_Callback(hObject, eventdata, handles)
% hObject    handle to edit14 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit14 as text
%        str2double(get(hObject,'String')) returns contents of edit14 as a double


% --- Executes during object creation, after setting all properties.
function edit14_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit14 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end


% --- Executes when figure1 is resized.
function figure1_SizeChangedFcn(hObject, eventdata, handles)
% hObject    handle to figure1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on slider movement.
function slider1_Callback(hObject, eventdata, handles)
% hObject    handle to slider1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
xyz = str2num(get(handles.cableDistance, 'String'))*get(handles.slider1, 'Value');
set(handles.text24, 'String', num2str(xyz));

% Hints: get(hObject,'Value') returns position of slider
%        get(hObject,'Min') and get(hObject,'Max') to determine range of slider


% --- Executes during object creation, after setting all properties.
function slider1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to slider1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: slider controls usually have a light gray background.
if isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor',[.9 .9 .9]);
end


% --- Executes on slider movement.
function slider2_Callback(hObject, eventdata, handles)
xyz = str2num(get(handles.bundlesDistance, 'String'))*get(handles.slider2, 'Value');
set(handles.text27, 'String', num2str(xyz));
% hObject    handle to slider2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'Value') returns position of slider
%        get(hObject,'Min') and get(hObject,'Max') to determine range of slider


% --- Executes during object creation, after setting all properties.
function slider2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to slider2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: slider controls usually have a light gray background.
if isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor',[.9 .9 .9]);
end



function PF_Field_Callback(hObject, eventdata, handles)
pf_Variable = get(handles.PF_Field, 'String');
if ((isempty(str2num(pf_Variable))) || (abs(str2num(pf_Variable)) > 1))
    set(handles.PF_Error_Text, 'String', 'Please write a valid value')
    assignin('base', 'Error_State', 1)
else
    set(handles.PF_Error_Text, 'String', '')
    assignin('base', 'Error_State', 0)
    assignin('base', 'PF', str2num(pf_Variable))
end
% hObject    handle to PF_Field (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of PF_Field as text
%        str2double(get(hObject,'String')) returns contents of PF_Field as a double


% --- Executes during object creation, after setting all properties.
function PF_Field_CreateFcn(hObject, eventdata, handles)
% hObject    handle to PF_Field (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end


% --- Executes during object creation, after setting all properties.
function figure1_CreateFcn(hObject, eventdata, handles)
assignin('base', 'Error_State', 0);
% hObject    handle to figure1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called


% --- Executes on button press in twoBundles.
function twoBundles_Callback(hObject, eventdata, handles)
assignin('base', 'Bundles_Number', 2)
% hObject    handle to twoBundles (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of twoBundles


% --- Executes on button press in threeBundles.
function threeBundles_Callback(hObject, eventdata, handles)
assignin('base', 'Bundles_Number', 3)
% hObject    handle to threeBundles (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of threeBundles


% --- Executes on button press in fourBundles.
function fourBundles_Callback(hObject, eventdata, handles)
assignin('base', 'Bundles_Number', 4)
% hObject    handle to fourBundles (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of fourBundles


% --- Executes on button press in piModelRadioButton.
function piModelRadioButton_Callback(hObject, eventdata, handles)
assignin('base', 'Model', 1)
% hObject    handle to piModelRadioButton (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of piModelRadioButton


% --- Executes on button press in tModelRadioButton.
function tModelRadioButton_Callback(hObject, eventdata, handles)
assignin('base', 'Model', 2)
% hObject    handle to tModelRadioButton (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of tModelRadioButton


% --- Executes on button press in radiobutton9.
function radiobutton9_Callback(hObject, eventdata, handles)
assignin('base', 'Frequency', 50)
% hObject    handle to radiobutton9 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of radiobutton9


% --- Executes on button press in radiobutton10.
function radiobutton10_Callback(hObject, eventdata, handles)
assignin('base', 'Frequency', 60)
% hObject    handle to radiobutton10 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of radiobutton10


% --- Executes during object creation, after setting all properties.
function twoBundles_CreateFcn(hObject, eventdata, handles)
% hObject    handle to twoBundles (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called



function Circuits_Number_Field_Callback(hObject, eventdata, handles)
Circuits_Number_Variable = get(handles.Circuits_Number_Field, 'String');
if ((isempty(str2num(Circuits_Number_Variable))) || (str2num(Circuits_Number_Variable) < 0))
    set(handles.Circuits_Number_Error, 'String', 'Please write a valid value')
    assignin('base', 'Error_State', 1)
else
    set(handles.Circuits_Number_Error, 'String', '')
    assignin('base', 'Error_State', 0)
    assignin('base', 'Circuits_Number', str2num(Circuits_Number_Variable))
end
% hObject    handle to Circuits_Number_Field (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of Circuits_Number_Field as text
%        str2double(get(hObject,'String')) returns contents of Circuits_Number_Field as a double


% --- Executes during object creation, after setting all properties.
function Circuits_Number_Field_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Circuits_Number_Field (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end



function requiredEfficiency_Callback(hObject, eventdata, handles)
requiredEfficiency_Variable = get(handles.requiredEfficiency, 'String');
if ((isempty(str2num(requiredEfficiency_Variable))) || (str2num(requiredEfficiency_Variable) < 0))
    set(handles.Efficiency_Error_Text, 'String', 'Please write a valid value')
    assignin('base', 'Error_State', 1)
else
    set(handles.Efficiency_Error_Text, 'String', '')
    assignin('base', 'Error_State', 0)
    assignin('base', 'Required_Efficiency', str2num(requiredEfficiency_Variable))
end
% hObject    handle to requiredEfficiency (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of requiredEfficiency as text
%        str2double(get(hObject,'String')) returns contents of requiredEfficiency as a double


% --- Executes during object creation, after setting all properties.
function requiredEfficiency_CreateFcn(hObject, eventdata, handles)
% hObject    handle to requiredEfficiency (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end



function requiredVR_Callback(hObject, eventdata, handles)
requiredVR_Variable = get(handles.requiredVR, 'String');
if ((isempty(str2num(requiredVR_Variable))) || (str2num(requiredVR_Variable) < 0))
    set(handles.VR_Error_Text, 'String', 'Please write a valid value')
    assignin('base', 'Error_State', 1)
else
    set(handles.VR_Error_Text, 'String', '')
    assignin('base', 'Error_State', 0)
    assignin('base', 'Required_VR', str2num(requiredVR_Variable))
end
% hObject    handle to requiredVR (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of requiredVR as text
%        str2double(get(hObject,'String')) returns contents of requiredVR as a double


% --- Executes during object creation, after setting all properties.
function requiredVR_CreateFcn(hObject, eventdata, handles)
% hObject    handle to requiredVR (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor', 'white');
end

% --- Executes during object creation, after setting all properties.
function Power_Error_Text_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Power_Error_Text (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called


% --- Executes during object creation, after setting all properties.
function noCable_CreateFcn(hObject, eventdata, handles)
% hObject    handle to noCable (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
