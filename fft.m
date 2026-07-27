%% ========================================================================
% 회전체 가속도 및 두 베어링 합산 반력 FFT 분석
%
% Excel 시트 이름:
%   0.5, 1, 2, 4, 6, 8, ...
%
% 목표 회전수:
%   목표 RPM = 숫자형 시트 이름 × 1e4
%
% 예:
%   시트 "0.5" -> 5,000 rpm
%   시트 "1"   -> 10,000 rpm
%   시트 "4"   -> 40,000 rpm
%
% Excel 열 구성:
%   A열 : 시간 [s]
%   B열 : X축 가속도 [m/s^2]
%   C열 : Y축 가속도 [m/s^2]
%   D열 : 지름 1 mm 위치의 접선방향 선속도 [m/s]
%   E열 : 중간+오른쪽 베어링의 합산 반력 Fx [N]
%   F열 : 중간+오른쪽 베어링의 합산 반력 Fy [N]
%   G열 : 중간+오른쪽 베어링의 합산 반력 Fz [N]
%
% 베어링 조건:
%   왼쪽 베어링:
%       radial, tangential, axial 모두 fixed
%
%   중간 베어링:
%       tangential free
%
%   오른쪽 끝 베어링:
%       tangential free
%
% 주의:
%   E:G열은 중간 베어링과 오른쪽 끝 베어링의 합산 반력입니다.
%
%   Fx_sum = Fx_middle + Fx_right
%   Fy_sum = Fy_middle + Fy_right
%   Fz_sum = Fz_middle + Fz_right
%
%   따라서 각 베어링의 개별 반력은 계산하지 않습니다.
%
% 단위:
%   길이       : m
%   시간       : s
%   선속도     : m/s
%   가속도     : m/s^2
%   힘         : N
%   각속도     : rad/s
%   회전속도   : rpm
%   주파수     : Hz
%
% 회전축:
%   전역 Z축
%
% D열 속도 기준:
%   지름   = 1 mm
%   반지름 = 0.5 mm = 0.0005 m
%
%   omega = v / r
%
% 기본 설정에서는 회전각과 고조파 주파수를 시트명의 RPM으로 계산합니다.
% D열은 목표 선속도와 일치하는지 검증하는 용도로 사용합니다.
%% ========================================================================

clear;
clc;
close all;


%% ========================================================================
% 1. 사용자 설정
%% ========================================================================

% ------------------------------------------------------------------------
% 입력 Excel 파일 경로
% ------------------------------------------------------------------------
% 직접 입력 예:
%
% inputFile = ...
%     'C:\Users\Administrator\Desktop\과제\FFT\FFT_raw.xlsx';
%
% 빈 문자열로 두면 파일 선택창이 열립니다.

inputFile = '';


% ------------------------------------------------------------------------
% 출력 폴더
% ------------------------------------------------------------------------
% 직접 입력 예:
%
% outputFolder = ...
%     'C:\Users\Administrator\Desktop\과제\FFT\MATLAB_Result';
%
% 빈 문자열로 두면 입력 파일 위치에 자동 생성됩니다.

outputFolder = '';


% ------------------------------------------------------------------------
% 회전속도 계산 기준
% ------------------------------------------------------------------------
% "sheetRPM":
%   시트 이름으로부터 회전속도를 계산합니다.
%   예: 시트 "4"이면 40,000 rpm을 사용합니다.
%
% "columnD":
%   D열 선속도와 반지름으로부터 회전속도를 계산합니다.
%
% 현재 D열 값이 목표 RPM과 맞지 않았으므로 "sheetRPM"을 권장합니다.

rotationSpeedSource = "sheetRPM";


% ------------------------------------------------------------------------
% D열 선속도가 정의된 지름
% ------------------------------------------------------------------------

rotationDiameter_m = 1.0e-3;       % 1 mm
rotationRadius_m = rotationDiameter_m / 2;


% ------------------------------------------------------------------------
% 초기 회전각
% ------------------------------------------------------------------------
% 회전각 0도에서 다음과 같이 가정합니다.
%
% radial 방향     = 전역 +X 방향
% tangential 방향 = 전역 +Y 방향
%
% 실제 초기 각도를 알고 있으면 수정하십시오.

initialRotationAngle_deg = 0;


% ------------------------------------------------------------------------
% FFT 분석 구간
% ------------------------------------------------------------------------
% 각 시트의 마지막 몇 회전을 분석할지 설정합니다.

analysisRevolutions = 30;


% FFT 분석에 필요한 최소 데이터 개수

minimumAnalysisPoints = 256;


% ------------------------------------------------------------------------
% FFT 설정
% ------------------------------------------------------------------------

maxPlotFrequency_Hz = 10500;

harmonicOrders = [1, 2, 3];

harmonicSearchFraction = 0.03;
minimumSearchHalfWidth_Hz = 5;

useHannWindow = true;
useZeroPadding = true;


% ------------------------------------------------------------------------
% 결과 저장 설정
% ------------------------------------------------------------------------

% 균일 시간 간격으로 보간된 시간 데이터를 Excel에 저장할지 설정
% 데이터가 매우 많으면 결과 파일 용량이 커집니다.

saveProcessedTimeHistory = false;


% 각 시트의 FFT 전체 데이터를 Excel에 저장할지 설정

saveSpectrumData = true;


% 그래프 화면 표시 여부
%
% 'on'  : MATLAB 화면에 표시
% 'off' : 파일로만 저장

figureVisibility = 'on';


% 그래프 저장 후 자동으로 닫을지 설정

closeFiguresAfterSaving = false;


%% ========================================================================
% 2. 입력 파일 선택 및 출력 폴더 생성
%% ========================================================================

if isempty(inputFile) || ~isfile(inputFile)

    [selectedFile, selectedPath] = uigetfile( ...
        {'*.xlsx;*.xls', 'Excel 파일 (*.xlsx, *.xls)'}, ...
        '분석할 Excel 파일을 선택하십시오.');

    if isequal(selectedFile, 0)
        error('파일 선택이 취소되었습니다.');
    end

    inputFile = fullfile(selectedPath, selectedFile);
end


[inputFolder, inputBaseName, ~] = fileparts(inputFile);


if isempty(outputFolder)

    outputFolder = fullfile( ...
        inputFolder, ...
        [inputBaseName, '_MATLAB_Result']);
end


if ~exist(outputFolder, 'dir')
    mkdir(outputFolder);
end


figureFolder = fullfile(outputFolder, 'Figures');

if ~exist(figureFolder, 'dir')
    mkdir(figureFolder);
end


timeStamp = datestr(now, 'yyyymmdd_HHMMSS');


outputExcel = fullfile( ...
    outputFolder, ...
    ['Rotor_Analysis_', timeStamp, '.xlsx']);


fprintf('\n');
fprintf('============================================================\n');

fprintf('입력 Excel 파일\n');
fprintf('%s\n\n', inputFile);

fprintf('회전속도 계산 기준      : %s\n', ...
    char(rotationSpeedSource));

fprintf('D열 속도 기준 지름      : %.6e m\n', ...
    rotationDiameter_m);

fprintf('D열 속도 기준 반지름    : %.6e m\n', ...
    rotationRadius_m);

fprintf('\n출력 폴더\n');
fprintf('%s\n\n', outputFolder);

fprintf('결과 Excel 파일\n');
fprintf('%s\n', outputExcel);

fprintf('============================================================\n\n');


%% ========================================================================
% 3. 숫자형 시트 검색
%% ========================================================================

allSheetNames = sheetnames(inputFile);

numberOfAllSheets = numel(allSheetNames);

sheetFactors = nan(numberOfAllSheets, 1);


for iSheet = 1:numberOfAllSheets

    sheetFactors(iSheet) = ...
        str2double(strtrim(allSheetNames(iSheet)));
end


% 숫자로 변환할 수 있는 시트만 선택합니다.

validSheetMask = ~isnan(sheetFactors);

sheetNames = allSheetNames(validSheetMask);
sheetFactors = sheetFactors(validSheetMask);


if isempty(sheetNames)

    error([ ...
        '숫자형 시트를 찾지 못했습니다. ', ...
        '시트 이름을 0.5, 1, 2, 4와 같이 설정하십시오.']);
end


% 숫자값 기준 오름차순으로 시트를 정렬합니다.

[sheetFactors, sortIndex] = sort(sheetFactors);

sheetNames = sheetNames(sortIndex);


fprintf('분석 대상 시트\n');

for iSheet = 1:numel(sheetNames)

    targetRPM = sheetFactors(iSheet) * 1e4;

    fprintf('  시트 %-12s -> %12.3f rpm\n', ...
        char(sheetNames(iSheet)), ...
        targetRPM);
end

fprintf('\n');


%% ========================================================================
% 4. 전체 시트 분석
%% ========================================================================

summaryRows = struct([]);


for iSheet = 1:numel(sheetNames)

    currentSheet = sheetNames(iSheet);

    sheetFactor = sheetFactors(iSheet);

    targetRPM = sheetFactor * 1e4;


    fprintf('------------------------------------------------------------\n');

    fprintf('[%d/%d] 시트 "%s" 분석 시작\n', ...
        iSheet, ...
        numel(sheetNames), ...
        char(currentSheet));

    fprintf('목표 회전수: %.3f rpm\n', targetRPM);


    %% --------------------------------------------------------------------
    % 4-1. Excel 데이터 읽기
    %% --------------------------------------------------------------------

    rawData = readmatrix( ...
        inputFile, ...
        'Sheet', char(currentSheet), ...
        'Range', 'A:G');


    if size(rawData, 2) < 7

        warning([ ...
            '시트 "%s"의 열 개수가 7개보다 적습니다. ', ...
            'A:G열이 필요하므로 해당 시트를 건너뜁니다.'], ...
            char(currentSheet));

        continue;
    end


    time_s = rawData(:, 1);

    accelerationX_m_s2 = rawData(:, 2);
    accelerationY_m_s2 = rawData(:, 3);

    columnDTangentialVelocity_m_s = rawData(:, 4);

    combinedBearingForceX_N = rawData(:, 5);
    combinedBearingForceY_N = rawData(:, 6);
    combinedBearingForceZ_N = rawData(:, 7);


    %% --------------------------------------------------------------------
    % 4-2. 시간 데이터가 유효한 행만 선택
    %% --------------------------------------------------------------------

    validTimeMask = isfinite(time_s);


    time_s = time_s(validTimeMask);

    accelerationX_m_s2 = ...
        accelerationX_m_s2(validTimeMask);

    accelerationY_m_s2 = ...
        accelerationY_m_s2(validTimeMask);

    columnDTangentialVelocity_m_s = ...
        columnDTangentialVelocity_m_s(validTimeMask);

    combinedBearingForceX_N = ...
        combinedBearingForceX_N(validTimeMask);

    combinedBearingForceY_N = ...
        combinedBearingForceY_N(validTimeMask);

    combinedBearingForceZ_N = ...
        combinedBearingForceZ_N(validTimeMask);


    if numel(time_s) < minimumAnalysisPoints

        warning([ ...
            '시트 "%s"의 유효 데이터가 %d개로 너무 적습니다. ', ...
            '해당 시트를 건너뜁니다.'], ...
            char(currentSheet), ...
            numel(time_s));

        continue;
    end


    %% --------------------------------------------------------------------
    % 4-3. 시간 순서 정렬
    %% --------------------------------------------------------------------

    [time_s, timeSortIndex] = sort(time_s);


    accelerationX_m_s2 = ...
        accelerationX_m_s2(timeSortIndex);

    accelerationY_m_s2 = ...
        accelerationY_m_s2(timeSortIndex);

    columnDTangentialVelocity_m_s = ...
        columnDTangentialVelocity_m_s(timeSortIndex);

    combinedBearingForceX_N = ...
        combinedBearingForceX_N(timeSortIndex);

    combinedBearingForceY_N = ...
        combinedBearingForceY_N(timeSortIndex);

    combinedBearingForceZ_N = ...
        combinedBearingForceZ_N(timeSortIndex);


    %% --------------------------------------------------------------------
    % 4-4. 중복 시간 제거
    %% --------------------------------------------------------------------

    [time_s, uniqueTimeIndex] = unique(time_s, 'stable');


    accelerationX_m_s2 = ...
        accelerationX_m_s2(uniqueTimeIndex);

    accelerationY_m_s2 = ...
        accelerationY_m_s2(uniqueTimeIndex);

    columnDTangentialVelocity_m_s = ...
        columnDTangentialVelocity_m_s(uniqueTimeIndex);

    combinedBearingForceX_N = ...
        combinedBearingForceX_N(uniqueTimeIndex);

    combinedBearingForceY_N = ...
        combinedBearingForceY_N(uniqueTimeIndex);

    combinedBearingForceZ_N = ...
        combinedBearingForceZ_N(uniqueTimeIndex);


    %% --------------------------------------------------------------------
    % 4-5. NaN 및 Inf 처리
    %% --------------------------------------------------------------------

    accelerationX_m_s2 = fillInvalidData( ...
        accelerationX_m_s2, ...
        'X축 가속도');

    accelerationY_m_s2 = fillInvalidData( ...
        accelerationY_m_s2, ...
        'Y축 가속도');

    columnDTangentialVelocity_m_s = fillInvalidData( ...
        columnDTangentialVelocity_m_s, ...
        'D열 접선 선속도');

    combinedBearingForceX_N = fillInvalidData( ...
        combinedBearingForceX_N, ...
        '두 베어링 합산 Fx');

    combinedBearingForceY_N = fillInvalidData( ...
        combinedBearingForceY_N, ...
        '두 베어링 합산 Fy');

    combinedBearingForceZ_N = fillInvalidData( ...
        combinedBearingForceZ_N, ...
        '두 베어링 합산 Fz');


    if any(diff(time_s) <= 0)

        error( ...
            '시트 "%s"의 시간 데이터가 단조 증가하지 않습니다.', ...
            char(currentSheet));
    end


    %% --------------------------------------------------------------------
    % 4-6. 목표 RPM에 해당하는 이론값 계산
    %% --------------------------------------------------------------------

    targetAngularVelocity_rad_s = ...
        targetRPM * 2*pi / 60;


    expectedTangentialVelocity_m_s = ...
        targetAngularVelocity_rad_s * rotationRadius_m;


    %% --------------------------------------------------------------------
    % 4-7. D열 선속도를 각속도와 RPM으로 변환
    %
    % omega_D = v_D / r
    %% --------------------------------------------------------------------

    columnDConvertedAngularVelocity_rad_s = ...
        columnDTangentialVelocity_m_s / rotationRadius_m;


    columnDConvertedRPM = ...
        columnDConvertedAngularVelocity_rad_s * 60 / (2*pi);


    %% --------------------------------------------------------------------
    % 4-8. 실제 분석에 사용할 회전속도 결정
    %% --------------------------------------------------------------------

    switch lower(string(rotationSpeedSource))

        case "sheetrpm"

            % 시트 이름의 목표 RPM을 전 구간에 일정하게 사용합니다.

            usedAngularVelocity_rad_s = ...
                targetAngularVelocity_rad_s * ones(size(time_s));


            usedRPM = ...
                targetRPM * ones(size(time_s));


        case "columnd"

            % D열의 선속도로부터 계산한 회전속도를 사용합니다.

            usedAngularVelocity_rad_s = ...
                columnDConvertedAngularVelocity_rad_s;


            usedRPM = ...
                columnDConvertedRPM;


        otherwise

            error([ ...
                'rotationSpeedSource는 "sheetRPM" 또는 ', ...
                '"columnD"이어야 합니다.']);
    end


    %% --------------------------------------------------------------------
    % 4-9. 회전각 계산
    %
    % theta(t) = theta(0) + integral(omega dt)
    %% --------------------------------------------------------------------

    initialRotationAngle_rad = ...
        deg2rad(initialRotationAngle_deg);


    rotationAngle_rad = ...
        initialRotationAngle_rad + ...
        cumtrapz(time_s, usedAngularVelocity_rad_s);


    %% --------------------------------------------------------------------
    % 4-10. 누적 회전 수 계산
    %% --------------------------------------------------------------------

    cumulativeRevolutions = ...
        cumtrapz( ...
        time_s, ...
        abs(usedAngularVelocity_rad_s)) / (2*pi);


    totalRevolutions = cumulativeRevolutions(end);


    %% --------------------------------------------------------------------
    % 4-11. 마지막 지정 회전 수에 해당하는 분석 구간 선택
    %% --------------------------------------------------------------------

    if totalRevolutions > analysisRevolutions

        analysisStartRevolution = ...
            totalRevolutions - analysisRevolutions;


        analysisMask = ...
            cumulativeRevolutions >= analysisStartRevolution;

    else

        analysisMask = true(size(time_s));


        warning([ ...
            '시트 "%s"의 전체 회전 수가 %.3f회로, ', ...
            '설정한 %.3f회보다 적습니다. 전체 구간을 사용합니다.'], ...
            char(currentSheet), ...
            totalRevolutions, ...
            analysisRevolutions);
    end


    % 분석 구간의 데이터 개수가 너무 적으면 마지막 데이터 구간 사용

    if nnz(analysisMask) < minimumAnalysisPoints

        startIndex = max( ...
            1, ...
            numel(time_s) - minimumAnalysisPoints + 1);


        analysisMask = false(size(time_s));

        analysisMask(startIndex:end) = true;
    end


    analysisTimeRaw_s = time_s(analysisMask);


    if numel(analysisTimeRaw_s) < 2

        warning( ...
            '시트 "%s"의 분석 시간 데이터가 부족합니다.', ...
            char(currentSheet));

        continue;
    end


    %% --------------------------------------------------------------------
    % 4-12. 균일 시간 간격 생성
    %% --------------------------------------------------------------------

    measuredTimeStep_s = ...
        median(diff(analysisTimeRaw_s));


    if ~isfinite(measuredTimeStep_s) || measuredTimeStep_s <= 0

        warning( ...
            '시트 "%s"의 시간 간격이 올바르지 않습니다.', ...
            char(currentSheet));

        continue;
    end


    uniformTime_s = ...
        ( ...
        analysisTimeRaw_s(1): ...
        measuredTimeStep_s: ...
        analysisTimeRaw_s(end) ...
        ).';


    if numel(uniformTime_s) < minimumAnalysisPoints

        warning([ ...
            '시트 "%s"의 균일 보간 후 데이터가 %d개로 ', ...
            '너무 적습니다.'], ...
            char(currentSheet), ...
            numel(uniformTime_s));

        continue;
    end


    %% --------------------------------------------------------------------
    % 4-13. 균일 시간 간격으로 모든 신호 보간
    %% --------------------------------------------------------------------

    columnDTangentialVelocityUniform_m_s = interpolateSignal( ...
        analysisTimeRaw_s, ...
        columnDTangentialVelocity_m_s(analysisMask), ...
        uniformTime_s);


    columnDConvertedRPMUniform = interpolateSignal( ...
        analysisTimeRaw_s, ...
        columnDConvertedRPM(analysisMask), ...
        uniformTime_s);


    usedAngularVelocityUniform_rad_s = interpolateSignal( ...
        analysisTimeRaw_s, ...
        usedAngularVelocity_rad_s(analysisMask), ...
        uniformTime_s);


    usedRPMUniform = interpolateSignal( ...
        analysisTimeRaw_s, ...
        usedRPM(analysisMask), ...
        uniformTime_s);


    rotationAngleUniform_rad = interpolateSignal( ...
        analysisTimeRaw_s, ...
        rotationAngle_rad(analysisMask), ...
        uniformTime_s);


    accelerationXUniform_m_s2 = interpolateSignal( ...
        analysisTimeRaw_s, ...
        accelerationX_m_s2(analysisMask), ...
        uniformTime_s);


    accelerationYUniform_m_s2 = interpolateSignal( ...
        analysisTimeRaw_s, ...
        accelerationY_m_s2(analysisMask), ...
        uniformTime_s);


    combinedBearingForceXUniform_N = interpolateSignal( ...
        analysisTimeRaw_s, ...
        combinedBearingForceX_N(analysisMask), ...
        uniformTime_s);


    combinedBearingForceYUniform_N = interpolateSignal( ...
        analysisTimeRaw_s, ...
        combinedBearingForceY_N(analysisMask), ...
        uniformTime_s);


    combinedBearingForceZUniform_N = interpolateSignal( ...
        analysisTimeRaw_s, ...
        combinedBearingForceZ_N(analysisMask), ...
        uniformTime_s);


    %% --------------------------------------------------------------------
    % 4-14. 전역 성분을 회전체 radial-tangential 성분으로 변환
    %
    % radial 단위벡터:
    %   e_r = [cos(theta), sin(theta)]
    %
    % tangential 단위벡터:
    %   e_t = [-sin(theta), cos(theta)]
    %
    % radial:
    %   F_r = Fx*cos(theta) + Fy*sin(theta)
    %
    % tangential:
    %   F_t = -Fx*sin(theta) + Fy*cos(theta)
    %% --------------------------------------------------------------------

    cosThetaUniform = cos(rotationAngleUniform