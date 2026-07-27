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

    cosThetaUniform = cos(rotationAngleUniform_rad);
    sinThetaUniform = sin(rotationAngleUniform_rad);


    % 가속도 radial 성분

    accelerationRadialUniform_m_s2 = ...
        accelerationXUniform_m_s2 .* cosThetaUniform + ...
        accelerationYUniform_m_s2 .* sinThetaUniform;


    % 가속도 tangential 성분

    accelerationTangentialUniform_m_s2 = ...
       -accelerationXUniform_m_s2 .* sinThetaUniform + ...
        accelerationYUniform_m_s2 .* cosThetaUniform;


    % 가속도 XY 크기

    accelerationXYMagnitudeUniform_m_s2 = ...
        sqrt( ...
        accelerationXUniform_m_s2.^2 + ...
        accelerationYUniform_m_s2.^2);


    % 합산 반력 radial 성분

    combinedBearingForceRadialUniform_N = ...
        combinedBearingForceXUniform_N .* cosThetaUniform + ...
        combinedBearingForceYUniform_N .* sinThetaUniform;


    % 합산 반력 tangential 성분

    combinedBearingForceTangentialUniform_N = ...
       -combinedBearingForceXUniform_N .* sinThetaUniform + ...
        combinedBearingForceYUniform_N .* cosThetaUniform;


    % 합산 반력 axial 성분

    combinedBearingForceAxialUniform_N = ...
        combinedBearingForceZUniform_N;


    % 합산 반력 XY 크기

    combinedBearingForceXYMagnitudeUniform_N = ...
        sqrt( ...
        combinedBearingForceXUniform_N.^2 + ...
        combinedBearingForceYUniform_N.^2);


    % 합산 반력 XYZ 크기

    combinedBearingForceXYZMagnitudeUniform_N = ...
        sqrt( ...
        combinedBearingForceXUniform_N.^2 + ...
        combinedBearingForceYUniform_N.^2 + ...
        combinedBearingForceZUniform_N.^2);


    %% --------------------------------------------------------------------
    % 4-15. 샘플링 및 회전속도 정보
    %% --------------------------------------------------------------------

    samplingFrequency_Hz = ...
        1 / measuredTimeStep_s;


    nyquistFrequency_Hz = ...
        samplingFrequency_Hz / 2;


    numberOfAnalysisSamples = ...
        numel(uniformTime_s);


    analysisDuration_s = ...
        uniformTime_s(end) - uniformTime_s(1);


    actualFrequencyResolution_Hz = ...
        samplingFrequency_Hz / numberOfAnalysisSamples;


    meanUsedAngularVelocity_rad_s = ...
        mean(usedAngularVelocityUniform_rad_s);


    meanAbsoluteUsedAngularVelocity_rad_s = ...
        mean(abs(usedAngularVelocityUniform_rad_s));


    meanUsedRPM = ...
        mean(usedRPMUniform);


    meanAbsoluteUsedRPM = ...
        mean(abs(usedRPMUniform));


    usedRotationFrequency_Hz = ...
        meanAbsoluteUsedRPM / 60;


    meanColumnDVelocity_m_s = ...
        mean(columnDTangentialVelocityUniform_m_s);


    meanAbsoluteColumnDVelocity_m_s = ...
        mean(abs(columnDTangentialVelocityUniform_m_s));


    meanAbsoluteColumnDConvertedRPM = ...
        mean(abs(columnDConvertedRPMUniform));


    if expectedTangentialVelocity_m_s ~= 0

        columnDVelocityErrorPercent = ...
            100 * ...
            ( ...
            meanAbsoluteColumnDVelocity_m_s - ...
            abs(expectedTangentialVelocity_m_s) ...
            ) / ...
            abs(expectedTangentialVelocity_m_s);

    else

        columnDVelocityErrorPercent = NaN;
    end


    firstAnalysisIndex = ...
        find(analysisMask, 1, 'first');


    analysisRevolutionsUsed = ...
        totalRevolutions - ...
        cumulativeRevolutions(firstAnalysisIndex);


    fprintf('목표 접선 선속도         : %.9f m/s\n', ...
        expectedTangentialVelocity_m_s);

    fprintf('D열 평균 접선 선속도     : %.9e m/s\n', ...
        meanColumnDVelocity_m_s);

    fprintf('D열 평균 절대 선속도     : %.9e m/s\n', ...
        meanAbsoluteColumnDVelocity_m_s);

    fprintf('D열 변환 평균 절대 RPM   : %.6f rpm\n', ...
        meanAbsoluteColumnDConvertedRPM);

    fprintf('분석에 사용한 평균 RPM   : %.6f rpm\n', ...
        meanAbsoluteUsedRPM);

    fprintf('분석에 사용한 평균 각속도: %.6f rad/s\n', ...
        meanAbsoluteUsedAngularVelocity_rad_s);

    fprintf('Sampling frequency       : %.6f Hz\n', ...
        samplingFrequency_Hz);

    fprintf('Nyquist frequency        : %.6f Hz\n', ...
        nyquistFrequency_Hz);

    fprintf('실제 주파수 분해능        : %.6f Hz\n', ...
        actualFrequencyResolution_Hz);

    fprintf('FFT 분석 데이터 수       : %d\n', ...
        numberOfAnalysisSamples);


    if isfinite(columnDVelocityErrorPercent) && ...
            abs(columnDVelocityErrorPercent) > 5

        warning([ ...
            '시트 "%s": D열 평균 절대 선속도가 목표 접선 선속도와 ', ...
            '%.3f%% 차이 납니다. 현재 회전각과 고조파 계산에는 ', ...
            '%s 기준을 사용합니다.'], ...
            char(currentSheet), ...
            columnDVelocityErrorPercent, ...
            char(rotationSpeedSource));
    end


    maximumRequiredHarmonicFrequency_Hz = ...
        max(harmonicOrders) * usedRotationFrequency_Hz;


    if nyquistFrequency_Hz < maximumRequiredHarmonicFrequency_Hz

        warning([ ...
            '시트 "%s"의 Nyquist 주파수 %.3f Hz가 ', ...
            '최대 분석 고조파 %.3f Hz보다 낮습니다.'], ...
            char(currentSheet), ...
            nyquistFrequency_Hz, ...
            maximumRequiredHarmonicFrequency_Hz);
    end


    %% --------------------------------------------------------------------
    % 4-16. 가속도 FFT 계산
    %% --------------------------------------------------------------------

    [frequency_Hz, accelerationXAmplitude_m_s2, fftLength] = ...
        singleSidedFFT( ...
        accelerationXUniform_m_s2, ...
        samplingFrequency_Hz, ...
        useHannWindow, ...
        useZeroPadding);


    [~, accelerationYAmplitude_m_s2] = ...
        singleSidedFFT( ...
        accelerationYUniform_m_s2, ...
        samplingFrequency_Hz, ...
        useHannWindow, ...
        useZeroPadding);


    [~, accelerationRadialAmplitude_m_s2] = ...
        singleSidedFFT( ...
        accelerationRadialUniform_m_s2, ...
        samplingFrequency_Hz, ...
        useHannWindow, ...
        useZeroPadding);


    [~, accelerationTangentialAmplitude_m_s2] = ...
        singleSidedFFT( ...
        accelerationTangentialUniform_m_s2, ...
        samplingFrequency_Hz, ...
        useHannWindow, ...
        useZeroPadding);


    % X와 Y 방향 FFT 진폭의 방향 합성값

    accelerationXYCombinedAmplitude_m_s2 = ...
        sqrt( ...
        accelerationXAmplitude_m_s2.^2 + ...
        accelerationYAmplitude_m_s2.^2);


    % Radial과 tangential 방향 FFT 진폭의 방향 합성값

    accelerationRotatingCombinedAmplitude_m_s2 = ...
        sqrt( ...
        accelerationRadialAmplitude_m_s2.^2 + ...
        accelerationTangentialAmplitude_m_s2.^2);


    %% --------------------------------------------------------------------
    % 4-17. 두 베어링 합산 반력 FFT 계산
    %% --------------------------------------------------------------------

    [~, combinedBearingForceXAmplitude_N] = ...
        singleSidedFFT( ...
        combinedBearingForceXUniform_N, ...
        samplingFrequency_Hz, ...
        useHannWindow, ...
        useZeroPadding);


    [~, combinedBearingForceYAmplitude_N] = ...
        singleSidedFFT( ...
        combinedBearingForceYUniform_N, ...
        samplingFrequency_Hz, ...
        useHannWindow, ...
        useZeroPadding);


    [~, combinedBearingForceZAmplitude_N] = ...
        singleSidedFFT( ...
        combinedBearingForceZUniform_N, ...
        samplingFrequency_Hz, ...
        useHannWindow, ...
        useZeroPadding);


    [~, combinedBearingForceRadialAmplitude_N] = ...
        singleSidedFFT( ...
        combinedBearingForceRadialUniform_N, ...
        samplingFrequency_Hz, ...
        useHannWindow, ...
        useZeroPadding);


    [~, combinedBearingForceTangentialAmplitude_N] = ...
        singleSidedFFT( ...
        combinedBearingForceTangentialUniform_N, ...
        samplingFrequency_Hz, ...
        useHannWindow, ...
        useZeroPadding);


    % Fx와 Fy FFT 진폭의 방향 합성값

    combinedBearingForceXYCombinedAmplitude_N = ...
        sqrt( ...
        combinedBearingForceXAmplitude_N.^2 + ...
        combinedBearingForceYAmplitude_N.^2);


    % Fx, Fy, Fz FFT 진폭의 방향 합성값

    combinedBearingForceXYZCombinedAmplitude_N = ...
        sqrt( ...
        combinedBearingForceXAmplitude_N.^2 + ...
        combinedBearingForceYAmplitude_N.^2 + ...
        combinedBearingForceZAmplitude_N.^2);


    % Radial과 tangential FFT 진폭의 방향 합성값

    combinedBearingForceRotatingCombinedAmplitude_N = ...
        sqrt( ...
        combinedBearingForceRadialAmplitude_N.^2 + ...
        combinedBearingForceTangentialAmplitude_N.^2);


    fftBinSpacing_Hz = ...
        samplingFrequency_Hz / fftLength;


    %% --------------------------------------------------------------------
    % 4-18. Summary 기본 데이터 생성
    %% --------------------------------------------------------------------

    summaryRow = struct();


    summaryRow.Sheet = ...
        string(currentSheet);


    summaryRow.TargetRPM = ...
        targetRPM;


    summaryRow.RotationSpeedSource = ...
        string(rotationSpeedSource);


    summaryRow.RotationDiameter_m = ...
        rotationDiameter_m;


    summaryRow.RotationRadius_m = ...
        rotationRadius_m;


    summaryRow.ExpectedTangentialVelocity_m_s = ...
        expectedTangentialVelocity_m_s;


    summaryRow.ColumnDMeanTangentialVelocity_m_s = ...
        meanColumnDVelocity_m_s;


    summaryRow.ColumnDMeanAbsoluteTangentialVelocity_m_s = ...
        meanAbsoluteColumnDVelocity_m_s;


    summaryRow.ColumnDConvertedMeanAbsoluteRPM = ...
        meanAbsoluteColumnDConvertedRPM;


    summaryRow.ColumnDVelocityErrorPercent = ...
        columnDVelocityErrorPercent;


    summaryRow.UsedMeanAngularVelocity_rad_s = ...
        meanUsedAngularVelocity_rad_s;


    summaryRow.UsedMeanAbsoluteAngularVelocity_rad_s = ...
        meanAbsoluteUsedAngularVelocity_rad_s;


    summaryRow.UsedMeanRPM = ...
        meanUsedRPM;


    summaryRow.UsedMeanAbsoluteRPM = ...
        meanAbsoluteUsedRPM;


    summaryRow.RotationFrequency_Hz = ...
        usedRotationFrequency_Hz;


    summaryRow.AnalysisStartTime_s = ...
        uniformTime_s(1);


    summaryRow.AnalysisEndTime_s = ...
        uniformTime_s(end);


    summaryRow.AnalysisDuration_s = ...
        analysisDuration_s;


    summaryRow.TotalRevolutions = ...
        totalRevolutions;


    summaryRow.AnalysisRevolutionsUsed = ...
        analysisRevolutionsUsed;


    summaryRow.NumberOfSamples = ...
        numberOfAnalysisSamples;


    summaryRow.TimeStep_s = ...
        measuredTimeStep_s;


    summaryRow.SamplingFrequency_Hz = ...
        samplingFrequency_Hz;


    summaryRow.NyquistFrequency_Hz = ...
        nyquistFrequency_Hz;


    summaryRow.ActualFrequencyResolution_Hz = ...
        actualFrequencyResolution_Hz;


    summaryRow.FFTBinSpacing_Hz = ...
        fftBinSpacing_Hz;


    %% --------------------------------------------------------------------
    % 4-19. 시간응답 통계 저장
    %% --------------------------------------------------------------------

    summaryRow.AccelerationXMean_m_s2 = ...
        mean(accelerationXUniform_m_s2);


    summaryRow.AccelerationYMean_m_s2 = ...
        mean(accelerationYUniform_m_s2);


    summaryRow.AccelerationXPeakAbsolute_m_s2 = ...
        max(abs(accelerationXUniform_m_s2));


    summaryRow.AccelerationYPeakAbsolute_m_s2 = ...
        max(abs(accelerationYUniform_m_s2));


    summaryRow.AccelerationRadialPeakAbsolute_m_s2 = ...
        max(abs(accelerationRadialUniform_m_s2));


    summaryRow.AccelerationTangentialPeakAbsolute_m_s2 = ...
        max(abs(accelerationTangentialUniform_m_s2));


    summaryRow.AccelerationXYMagnitudePeak_m_s2 = ...
        max(accelerationXYMagnitudeUniform_m_s2);


    summaryRow.CombinedBearingForceXMean_N = ...
        mean(combinedBearingForceXUniform_N);


    summaryRow.CombinedBearingForceYMean_N = ...
        mean(combinedBearingForceYUniform_N);


    summaryRow.CombinedBearingForceZMean_N = ...
        mean(combinedBearingForceZUniform_N);


    summaryRow.CombinedBearingForceXPeakAbsolute_N = ...
        max(abs(combinedBearingForceXUniform_N));


    summaryRow.CombinedBearingForceYPeakAbsolute_N = ...
        max(abs(combinedBearingForceYUniform_N));


    summaryRow.CombinedBearingForceZPeakAbsolute_N = ...
        max(abs(combinedBearingForceZUniform_N));


    summaryRow.CombinedBearingRadialPeakAbsolute_N = ...
        max(abs(combinedBearingForceRadialUniform_N));


    summaryRow.CombinedBearingTangentialPeakAbsolute_N = ...
        max(abs(combinedBearingForceTangentialUniform_N));


    summaryRow.CombinedBearingAxialPeakAbsolute_N = ...
        max(abs(combinedBearingForceAxialUniform_N));


    summaryRow.CombinedBearingXYMagnitudePeak_N = ...
        max(combinedBearingForceXYMagnitudeUniform_N);


    summaryRow.CombinedBearingXYZMagnitudePeak_N = ...
        max(combinedBearingForceXYZMagnitudeUniform_N);


    %% --------------------------------------------------------------------
    % 4-20. 1X, 2X, 3X 피크 추출
    %% --------------------------------------------------------------------

    for iOrder = 1:numel(harmonicOrders)

        currentOrder = harmonicOrders(iOrder);


        expectedFrequency_Hz = ...
            currentOrder * usedRotationFrequency_Hz;


        searchHalfWidth_Hz = max( ...
            expectedFrequency_Hz * harmonicSearchFraction, ...
            minimumSearchHalfWidth_Hz);


        % 가속도 XY 합성 진폭의 피크

        [accelerationPeakFrequency_Hz, ...
         accelerationPeakAmplitude_m_s2] = ...
            findPeakNearFrequency( ...
            frequency_Hz, ...
            accelerationXYCombinedAmplitude_m_s2, ...
            expectedFrequency_Hz, ...
            searchHalfWidth_Hz);


        % 합산 radial 반력 피크

        [radialForcePeakFrequency_Hz, ...
         radialForcePeakAmplitude_N] = ...
            findPeakNearFrequency( ...
            frequency_Hz, ...
            combinedBearingForceRadialAmplitude_N, ...
            expectedFrequency_Hz, ...
            searchHalfWidth_Hz);


        % 합산 tangential 반력 피크

        [tangentialForcePeakFrequency_Hz, ...
         tangentialForcePeakAmplitude_N] = ...
            findPeakNearFrequency( ...
            frequency_Hz, ...
            combinedBearingForceTangentialAmplitude_N, ...
            expectedFrequency_Hz, ...
            searchHalfWidth_Hz);


        % 합산 axial 반력 피크

        [axialForcePeakFrequency_Hz, ...
         axialForcePeakAmplitude_N] = ...
            findPeakNearFrequency( ...
            frequency_Hz, ...
            combinedBearingForceZAmplitude_N, ...
            expectedFrequency_Hz, ...
            searchHalfWidth_Hz);


        orderText = sprintf('%dX', currentOrder);


        % MATLAB 필드명은 숫자로 시작할 수 없으므로
        % Harmonic_1X, Harmonic_2X 등의 이름을 사용합니다.

        summaryRow.(sprintf( ...
            'Harmonic_%s_ExpectedFrequency_Hz', ...
            orderText)) = ...
            expectedFrequency_Hz;


        summaryRow.(sprintf( ...
            'AccelerationXY_%s_PeakFrequency_Hz', ...
            orderText)) = ...
            accelerationPeakFrequency_Hz;


        summaryRow.(sprintf( ...
            'AccelerationXY_%s_Amplitude_m_s2', ...
            orderText)) = ...
            accelerationPeakAmplitude_m_s2;


        summaryRow.(sprintf( ...
            'CombinedRadialForce_%s_PeakFrequency_Hz', ...
            orderText)) = ...
            radialForcePeakFrequency_Hz;


        summaryRow.(sprintf( ...
            'CombinedRadialForce_%s_Amplitude_N', ...
            orderText)) = ...
            radialForcePeakAmplitude_N;


        summaryRow.(sprintf( ...
            'CombinedTangentialForce_%s_PeakFrequency_Hz', ...
            orderText)) = ...
            tangentialForcePeakFrequency_Hz;


        summaryRow.(sprintf( ...
            'CombinedTangentialForce_%s_Amplitude_N', ...
            orderText)) = ...
            tangentialForcePeakAmplitude_N;


        summaryRow.(sprintf( ...
            'CombinedAxialForce_%s_PeakFrequency_Hz', ...
            orderText)) = ...
            axialForcePeakFrequency_Hz;


        summaryRow.(sprintf( ...
            'CombinedAxialForce_%s_Amplitude_N', ...
            orderText)) = ...
            axialForcePeakAmplitude_N;
    end


    if isempty(summaryRows)

        summaryRows = summaryRow;

    else

        summaryRows(end+1) = summaryRow;
    end


    %% --------------------------------------------------------------------
    % 4-21. 출력용 시트 이름 생성
    %% --------------------------------------------------------------------

    safeSheetText = regexprep( ...
        char(currentSheet), ...
        '[^a-zA-Z0-9]', ...
        '_');


    %% --------------------------------------------------------------------
    % 4-22. 보간된 시간응답 데이터 저장
    %% --------------------------------------------------------------------

    if saveProcessedTimeHistory

        processedTimeTable = table( ...
            uniformTime_s, ...
            columnDTangentialVelocityUniform_m_s, ...
            columnDConvertedRPMUniform, ...
            usedAngularVelocityUniform_rad_s, ...
            usedRPMUniform, ...
            rad2deg(rotationAngleUniform_rad), ...
            accelerationXUniform_m_s2, ...
            accelerationYUniform_m_s2, ...
            accelerationRadialUniform_m_s2, ...
            accelerationTangentialUniform_m_s2, ...
            accelerationXYMagnitudeUniform_m_s2, ...
            combinedBearingForceXUniform_N, ...
            combinedBearingForceYUniform_N, ...
            combinedBearingForceZUniform_N, ...
            combinedBearingForceRadialUniform_N, ...
            combinedBearingForceTangentialUniform_N, ...
            combinedBearingForceAxialUniform_N, ...
            combinedBearingForceXYMagnitudeUniform_N, ...
            combinedBearingForceXYZMagnitudeUniform_N, ...
            'VariableNames', { ...
            'Time_s', ...
            'ColumnD_TangentialVelocity_m_s', ...
            'ColumnD_ConvertedRPM', ...
            'UsedAngularVelocity_rad_s', ...
            'UsedRPM', ...
            'RotationAngle_deg', ...
            'AccelerationX_m_s2', ...
            'AccelerationY_m_s2', ...
            'AccelerationRadial_m_s2', ...
            'AccelerationTangential_m_s2', ...
            'AccelerationXYMagnitude_m_s2', ...
            'CombinedBearingForceX_N', ...
            'CombinedBearingForceY_N', ...
            'CombinedBearingForceZ_N', ...
            'CombinedBearingForceRadial_N', ...
            'CombinedBearingForceTangential_N', ...
            'CombinedBearingForceAxial_N', ...
            'CombinedBearingForceXYMagnitude_N', ...
            'CombinedBearingForceXYZMagnitude_N'});


        outputTimeSheet = makeExcelSheetName( ...
            'Time_', ...
            safeSheetText);


        writetable( ...
            processedTimeTable, ...
            outputExcel, ...
            'Sheet', outputTimeSheet);
    end


    %% --------------------------------------------------------------------
    % 4-23. FFT 데이터 저장
    %% --------------------------------------------------------------------

    if saveSpectrumData

        spectrumTable = table( ...
            frequency_Hz, ...
            accelerationXAmplitude_m_s2, ...
            accelerationYAmplitude_m_s2, ...
            accelerationXYCombinedAmplitude_m_s2, ...
            accelerationRadialAmplitude_m_s2, ...
            accelerationTangentialAmplitude_m_s2, ...
            accelerationRotatingCombinedAmplitude_m_s2, ...
            combinedBearingForceXAmplitude_N, ...
            combinedBearingForceYAmplitude_N, ...
            combinedBearingForceZAmplitude_N, ...
            combinedBearingForceXYCombinedAmplitude_N, ...
            combinedBearingForceXYZCombinedAmplitude_N, ...
            combinedBearingForceRadialAmplitude_N, ...
            combinedBearingForceTangentialAmplitude_N, ...
            combinedBearingForceRotatingCombinedAmplitude_N, ...
            'VariableNames', { ...
            'Frequency_Hz', ...
            'AccelerationX_Amplitude_m_s2', ...
            'AccelerationY_Amplitude_m_s2', ...
            'AccelerationXY_CombinedAmplitude_m_s2', ...
            'AccelerationRadial_Amplitude_m_s2', ...
            'AccelerationTangential_Amplitude_m_s2', ...
            'AccelerationRotating_CombinedAmplitude_m_s2', ...
            'CombinedBearingForceX_Amplitude_N', ...
            'CombinedBearingForceY_Amplitude_N', ...
            'CombinedBearingForceZ_Amplitude_N', ...
            'CombinedBearingForceXY_CombinedAmplitude_N', ...
            'CombinedBearingForceXYZ_CombinedAmplitude_N', ...
            'CombinedBearingForceRadial_Amplitude_N', ...
            'CombinedBearingForceTangential_Amplitude_N', ...
            'CombinedBearingForceRotating_CombinedAmplitude_N'});


        outputFFTSheet = makeExcelSheetName( ...
            'FFT_', ...
            safeSheetText);


        writetable( ...
            spectrumTable, ...
            outputExcel, ...
            'Sheet', outputFFTSheet);
    end


    %% --------------------------------------------------------------------
    % 4-24. 시간응답 그래프
    %% --------------------------------------------------------------------

    relativeTime_s = ...
        uniformTime_s - uniformTime_s(1);


    timeFigure = figure( ...
        'Visible', figureVisibility, ...
        'Color', 'w', ...
        'Position', [80, 40, 1500, 950]);


    tiledlayout(3, 2, ...
        'TileSpacing', 'compact', ...
        'Padding', 'compact');


    % D열 접선 선속도

    nexttile;

    plot( ...
        relativeTime_s, ...
        columnDTangentialVelocityUniform_m_s, ...
        'LineWidth', 1.0);

    hold on;

    yline( ...
        expectedTangentialVelocity_m_s, ...
        '--', ...
        sprintf('예상 %.6f m/s', ...
        expectedTangentialVelocity_m_s), ...
        'HandleVisibility', 'off');

    grid on;

    xlabel('시간 [s]');
    ylabel('접선 선속도 [m/s]');
    title('D열 접선 선속도');

    legend( ...
        'D열 측정값', ...
        'Location', 'best');


    % 분석 RPM과 D열 변환 RPM

    nexttile;

    plot( ...
        relativeTime_s, ...
        usedRPMUniform, ...
        'LineWidth', 1.0);

    hold on;

    plot( ...
        relativeTime_s, ...
        columnDConvertedRPMUniform, ...
        'LineWidth', 0.9);

    yline( ...
        targetRPM, ...
        '--', ...
        sprintf('목표 %.0f rpm', targetRPM), ...
        'HandleVisibility', 'off');

    grid on;

    xlabel('시간 [s]');
    ylabel('회전속도 [rpm]');
    title('분석 RPM과 D열 변환 RPM');

    legend( ...
        '분석에 사용한 RPM', ...
        'D열 변환 RPM', ...
        'Location', 'best');


    % 전역 가속도

    nexttile;

    plot( ...
        relativeTime_s, ...
        accelerationXUniform_m_s2, ...
        'LineWidth', 0.9);

    hold on;

    plot( ...
        relativeTime_s, ...
        accelerationYUniform_m_s2, ...
        'LineWidth', 0.9);

    plot( ...
        relativeTime_s, ...
        accelerationXYMagnitudeUniform_m_s2, ...
        'LineWidth', 1.0);

    grid on;

    xlabel('시간 [s]');
    ylabel('가속도 [m/s^2]');
    title('전역 좌표계 가속도');

    legend( ...
        'X', ...
        'Y', ...
        'XY 크기', ...
        'Location', 'best');


    % 회전체 좌표계 가속도

    nexttile;

    plot( ...
        relativeTime_s, ...
        accelerationRadialUniform_m_s2, ...
        'LineWidth', 0.9);

    hold on;

    plot( ...
        relativeTime_s, ...
        accelerationTangentialUniform_m_s2, ...
        'LineWidth', 0.9);

    grid on;

    xlabel('시간 [s]');
    ylabel('가속도 [m/s^2]');
    title('회전체 좌표계 가속도');

    legend( ...
        'Radial', ...
        'Tangential', ...
        'Location', 'best');


    % 합산 반력의 전역 성분

    nexttile;

    plot( ...
        relativeTime_s, ...
        combinedBearingForceXUniform_N, ...
        'LineWidth', 0.9);

    hold on;

    plot( ...
        relativeTime_s, ...
        combinedBearingForceYUniform_N, ...
        'LineWidth', 0.9);

    plot( ...
        relativeTime_s, ...
        combinedBearingForceZUniform_N, ...
        'LineWidth', 0.9);

    grid on;

    xlabel('시간 [s]');
    ylabel('합산 반력 [N]');
    title('두 베어링 합산 반력: 전역 성분');

    legend( ...
        'F_x', ...
        'F_y', ...
        'F_z', ...
        'Location', 'best');


    % 합산 반력의 회전체 좌표계 성분

    nexttile;

    plot( ...
        relativeTime_s, ...
        combinedBearingForceRadialUniform_N, ...
        'LineWidth', 0.9);

    hold on;

    plot( ...
        relativeTime_s, ...
        combinedBearingForceTangentialUniform_N, ...
        'LineWidth', 0.9);

    plot( ...
        relativeTime_s, ...
        combinedBearingForceAxialUniform_N, ...
        'LineWidth', 0.9);

    grid on;

    xlabel('시간 [s]');
    ylabel('합산 반력 [N]');
    title('두 베어링 합산 반력: 회전체 좌표계');

    legend( ...
        'Radial', ...
        'Tangential', ...
        'Axial', ...
        'Location', 'best');


    sgtitle(sprintf( ...
        '시트 %s | 목표 %.0f rpm | 분석 평균 %.1f rpm', ...
        char(currentSheet), ...
        targetRPM, ...
        meanAbsoluteUsedRPM));


    timeFigureFile = fullfile( ...
        figureFolder, ...
        sprintf( ...
        'Time_Sheet_%s_%drpm.png', ...
        safeSheetText, ...
        round(targetRPM)));


    exportgraphics( ...
        timeFigure, ...
        timeFigureFile, ...
        'Resolution', 200);


    %% --------------------------------------------------------------------
    % 4-25. FFT 그래프
    %% --------------------------------------------------------------------

    fftFigure = figure( ...
        'Visible', figureVisibility, ...
        'Color', 'w', ...
        'Position', [100, 50, 1500, 900]);


    tiledlayout(2, 2, ...
        'TileSpacing', 'compact', ...
        'Padding', 'compact');


    displayMaximumFrequency_Hz = min( ...
        maxPlotFrequency_Hz, ...
        nyquistFrequency_Hz);


    % 전역 가속도 FFT

    nexttile;

    semilogy( ...
        frequency_Hz, ...
        max(accelerationXAmplitude_m_s2, realmin), ...
        'LineWidth', 0.9);

    hold on;

    semilogy( ...
        frequency_Hz, ...
        max(accelerationYAmplitude_m_s2, realmin), ...
        'LineWidth', 0.9);

    semilogy( ...
        frequency_Hz, ...
        max(accelerationXYCombinedAmplitude_m_s2, realmin), ...
        'LineWidth', 1.0);

    addHarmonicLines( ...
        usedRotationFrequency_Hz, ...
        harmonicOrders);

    xlim([0, displayMaximumFrequency_Hz]);

    grid on;

    xlabel('주파수 [Hz]');
    ylabel('가속도 진폭 [m/s^2]');
    title('전역 가속도 FFT');

    legend( ...
        'X', ...
        'Y', ...
        'XY 합성', ...
        'Location', 'best');


    % 회전체 좌표계 가속도 FFT

    nexttile;

    semilogy( ...
        frequency_Hz, ...
        max(accelerationRadialAmplitude_m_s2, realmin), ...
        'LineWidth', 0.9);

    hold on;

    semilogy( ...
        frequency_Hz, ...
        max(accelerationTangentialAmplitude_m_s2, realmin), ...
        'LineWidth', 0.9);

    semilogy( ...
        frequency_Hz, ...
        max(accelerationRotatingCombinedAmplitude_m_s2, realmin), ...
        'LineWidth', 1.0);

    addHarmonicLines( ...
        usedRotationFrequency_Hz, ...
        harmonicOrders);

    xlim([0, displayMaximumFrequency_Hz]);

    grid on;

    xlabel('주파수 [Hz]');
    ylabel('가속도 진폭 [m/s^2]');
    title('회전체 좌표계 가속도 FFT');

    legend( ...
        'Radial', ...
        'Tangential', ...
        '방향 합성', ...
        'Location', 'best');


    % 합산 반력 전역 FFT

    nexttile;

    semilogy( ...
        frequency_Hz, ...
        max(combinedBearingForceXAmplitude_N, realmin), ...
        'LineWidth', 0.9);

    hold on;

    semilogy( ...
        frequency_Hz, ...
        max(combinedBearingForceYAmplitude_N, realmin), ...
        'LineWidth', 0.9);

    semilogy( ...
        frequency_Hz, ...
        max(combinedBearingForceZAmplitude_N, realmin), ...
        'LineWidth', 0.9);

    addHarmonicLines( ...
        usedRotationFrequency_Hz, ...
        harmonicOrders);

    xlim([0, displayMaximumFrequency_Hz]);

    grid on;

    xlabel('주파수 [Hz]');
    ylabel('합산 반력 진폭 [N]');
    title('두 베어링 합산 반력 FFT: 전역 성분');

    legend( ...
        'F_x', ...
        'F_y', ...
        'F_z', ...
        'Location', 'best');


    % 합산 반력 회전체 좌표계 FFT

    nexttile;

    semilogy( ...
        frequency_Hz, ...
        max(combinedBearingForceRadialAmplitude_N, realmin), ...
        'LineWidth', 0.9);

    hold on;

    semilogy( ...
        frequency_Hz, ...
        max(combinedBearingForceTangentialAmplitude_N, realmin), ...
        'LineWidth', 0.9);

    semilogy( ...
        frequency_Hz, ...
        max(combinedBearingForceZAmplitude_N, realmin), ...
        'LineWidth', 0.9);

    addHarmonicLines( ...
        usedRotationFrequency_Hz, ...
        harmonicOrders);

    xlim([0, displayMaximumFrequency_Hz]);

    grid on;

    xlabel('주파수 [Hz]');
    ylabel('합산 반력 진폭 [N]');
    title('두 베어링 합산 반력 FFT: 회전체 좌표계');

    legend( ...
        'Radial', ...
        'Tangential', ...
        'Axial', ...
        'Location', 'best');


    sgtitle(sprintf( ...
        'FFT | 시트 %s | 목표 %.0f rpm | 분석 평균 %.1f rpm', ...
        char(currentSheet), ...
        targetRPM, ...
        meanAbsoluteUsedRPM));


    fftFigureFile = fullfile( ...
        figureFolder, ...
        sprintf( ...
        'FFT_Sheet_%s_%drpm.png', ...
        safeSheetText, ...
        round(targetRPM)));


    exportgraphics( ...
        fftFigure, ...
        fftFigureFile, ...
        'Resolution', 200);


    if closeFiguresAfterSaving

        close(timeFigure);
        close(fftFigure);
    end


    fprintf('시트 "%s" 분석 완료\n\n', ...
        char(currentSheet));

end


%% ========================================================================
% 5. 전체 Summary 저장
%% ========================================================================

if isempty(summaryRows)

    error('분석에 성공한 숫자형 시트가 없습니다.');
end


summaryTable = struct2table(summaryRows);


summaryTable = sortrows( ...
    summaryTable, ...
    'TargetRPM');


writetable( ...
    summaryTable, ...
    outputExcel, ...
    'Sheet', 'Summary');


%% ========================================================================
% 6. 회전수별 비교 그래프
%% ========================================================================

summaryFigure = figure( ...
    'Visible', figureVisibility, ...
    'Color', 'w', ...
    'Position', [120, 40, 1500, 950]);


tiledlayout(3, 2, ...
    'TileSpacing', 'compact', ...
    'Padding', 'compact');


targetRPMAxis = ...
    summaryTable.TargetRPM;


% 목표 RPM과 분석 RPM 비교

nexttile;

plot( ...
    targetRPMAxis, ...
    summaryTable.UsedMeanAbsoluteRPM, ...
    '-o', ...
    'LineWidth', 1.2);

hold on;

plot( ...
    targetRPMAxis, ...
    targetRPMAxis, ...
    '--', ...
    'LineWidth', 1.0);

grid on;

xlabel('목표 회전속도 [rpm]');
ylabel('분석에 사용한 회전속도 [rpm]');
title('목표 RPM과 분석 RPM 비교');

legend( ...
    '분석 RPM', ...
    '목표 RPM', ...
    'Location', 'best');


% D열 평균 속도와 예상 속도 비교

nexttile;

plot( ...
    targetRPMAxis, ...
    summaryTable.ColumnDMeanAbsoluteTangentialVelocity_m_s, ...
    '-o', ...
    'LineWidth', 1.2);

hold on;

plot( ...
    targetRPMAxis, ...
    abs(summaryTable.ExpectedTangentialVelocity_m_s), ...
    '--o', ...
    'LineWidth', 1.0);

grid on;

xlabel('목표 회전속도 [rpm]');
ylabel('접선 선속도 [m/s]');
title('D열 속도와 예상 접선속도 비교');

legend( ...
    'D열 평균 절대속도', ...
    '예상 접선속도', ...
    'Location', 'best');


% 가속도 고조파

nexttile;

hold on;

for iOrder = 1:numel(harmonicOrders)

    currentOrder = harmonicOrders(iOrder);


    variableName = sprintf( ...
        'AccelerationXY_%dX_Amplitude_m_s2', ...
        currentOrder);


    plot( ...
        targetRPMAxis, ...
        summaryTable.(variableName), ...
        '-o', ...
        'LineWidth', 1.1, ...
        'DisplayName', sprintf('%dX', currentOrder));
end

grid on;

xlabel('목표 회전속도 [rpm]');
ylabel('가속도 진폭 [m/s^2]');
title('가속도 고조파 진폭');

legend('Location', 'best');


% 합산 radial 반력 고조파

nexttile;

hold on;

for iOrder = 1:numel(harmonicOrders)

    currentOrder = harmonicOrders(iOrder);


    variableName = sprintf( ...
        'CombinedRadialForce_%dX_Amplitude_N', ...
        currentOrder);


    plot( ...
        targetRPMAxis, ...
        summaryTable.(variableName), ...
        '-o', ...
        'LineWidth', 1.1, ...
        'DisplayName', sprintf('%dX', currentOrder));
end

grid on;

xlabel('목표 회전속도 [rpm]');
ylabel('합산 반력 진폭 [N]');
title('합산 radial 반력 고조파');

legend('Location', 'best');


% 합산 tangential 반력 고조파

nexttile;

hold on;

for iOrder = 1:numel(harmonicOrders)

    currentOrder = harmonicOrders(iOrder);


    variableName = sprintf( ...
        'CombinedTangentialForce_%dX_Amplitude_N', ...
        currentOrder);


    plot( ...
        targetRPMAxis, ...
        summaryTable.(variableName), ...
        '-o', ...
        'LineWidth', 1.1, ...
        'DisplayName', sprintf('%dX', currentOrder));
end

grid on;

xlabel('목표 회전속도 [rpm]');
ylabel('합산 반력 진폭 [N]');
title('합산 tangential 반력 고조파');

legend('Location', 'best');


% 합산 axial 반력 고조파

nexttile;

hold on;

for iOrder = 1:numel(harmonicOrders)

    currentOrder = harmonicOrders(iOrder);


    variableName = sprintf( ...
        'CombinedAxialForce_%dX_Amplitude_N', ...
        currentOrder);


    plot( ...
        targetRPMAxis, ...
        summaryTable.(variableName), ...
        '-o', ...
        'LineWidth', 1.1, ...
        'DisplayName', sprintf('%dX', currentOrder));
end

grid on;

xlabel('목표 회전속도 [rpm]');
ylabel('합산 반력 진폭 [N]');
title('합산 axial 반력 고조파');

legend('Location', 'best');


sgtitle( ...
    '회전체 가속도 및 두 베어링 합산 반력 분석');


summaryFigureFile = fullfile( ...
    figureFolder, ...
    ['Summary_', timeStamp, '.png']);


exportgraphics( ...
    summaryFigure, ...
    summaryFigureFile, ...
    'Resolution', 200);


if closeFiguresAfterSaving
    close(summaryFigure);
end


%% ========================================================================
% 7. 분석 완료 메시지
%% ========================================================================

fprintf('\n');
fprintf('============================================================\n');

fprintf('전체 분석이 완료되었습니다.\n\n');

fprintf('결과 Excel 파일\n');
fprintf('%s\n\n', outputExcel);

fprintf('그래프 저장 폴더\n');
fprintf('%s\n', figureFolder);

fprintf('============================================================\n');


%% ========================================================================
% 지역 함수
%% ========================================================================

function data = fillInvalidData(data, variableDescription)
% NaN 및 Inf 값을 선형 보간합니다.

    data = data(:);

    data(~isfinite(data)) = NaN;


    if all(isnan(data))

        error( ...
            '%s 데이터가 모두 NaN 또는 Inf입니다.', ...
            variableDescription);
    end


    data = fillmissing( ...
        data, ...
        'linear', ...
        'EndValues', 'nearest');
end


function outputSignal = interpolateSignal( ...
    originalTime, ...
    originalSignal, ...
    newTime)
% 원본 신호를 새로운 균일 시간축으로 선형 보간합니다.

    outputSignal = interp1( ...
        originalTime, ...
        originalSignal, ...
        newTime, ...
        'linear');
end


function [frequency_Hz, amplitude, fftLength] = singleSidedFFT( ...
    signal, ...
    samplingFrequency_Hz, ...
    useHannWindow, ...
    useZeroPadding)
% 단측 FFT 진폭 스펙트럼을 계산합니다.
%
% 출력 진폭은 정현파의 peak amplitude 기준입니다.
%
% 평균과 선형 추세를 제거한 뒤 FFT를 수행합니다.
% Hann window 사용 시 coherent gain을 보정합니다.

    signal = signal(:);

    numberOfSamples = numel(signal);


    if numberOfSamples < 4

        error('FFT 계산을 위한 데이터 개수가 너무 적습니다.');
    end


    % 평균 및 선형 추세 제거

    signal = detrend(signal, 'linear');


    % Hann window 생성

    if useHannWindow

        sampleIndex = (0:numberOfSamples-1).';


        window = ...
            0.5 - ...
            0.5*cos( ...
            2*pi*sampleIndex / numberOfSamples);

    else

        window = ones(numberOfSamples, 1);
    end


    windowedSignal = signal .* window;


    % FFT 길이 결정

    if useZeroPadding

        fftLength = ...
            2^nextpow2(numberOfSamples);

    else

        fftLength = ...
            numberOfSamples;
    end


    fftResult = fft( ...
        windowedSignal, ...
        fftLength);


    numberOfSingleSidedPoints = ...
        floor(fftLength / 2) + 1;


    % Hann window coherent gain 보정

    amplitude = ...
        2 * ...
        abs(fftResult(1:numberOfSingleSidedPoints)) / ...
        sum(window);


    % DC 성분은 두 배 하면 안 됨

    amplitude(1) = amplitude(1) / 2;


    % Nyquist 성분도 두 배 하면 안 됨

    if rem(fftLength, 2) == 0

        amplitude(end) = amplitude(end) / 2;
    end


    frequency_Hz = ...
        (0:numberOfSingleSidedPoints-1).' * ...
        samplingFrequency_Hz / fftLength;
end


function [peakFrequency_Hz, peakAmplitude] = ...
    findPeakNearFrequency( ...
    frequency_Hz, ...
    amplitude, ...
    expectedFrequency_Hz, ...
    searchHalfWidth_Hz)
% 예상 주파수 주변에서 가장 큰 FFT 진폭을 검색합니다.

    if ~isfinite(expectedFrequency_Hz) || ...
            expectedFrequency_Hz < 0

        peakFrequency_Hz = NaN;
        peakAmplitude = NaN;

        return;
    end


    searchMask = ...
        frequency_Hz >= ...
        max(0, expectedFrequency_Hz - searchHalfWidth_Hz) & ...
        frequency_Hz <= ...
        expectedFrequency_Hz + searchHalfWidth_Hz;


    if ~any(searchMask)

        peakFrequency_Hz = NaN;
        peakAmplitude = NaN;

        return;
    end


    searchFrequency_Hz = ...
        frequency_Hz(searchMask);


    searchAmplitude = ...
        amplitude(searchMask);


    [peakAmplitude, localPeakIndex] = ...
        max(searchAmplitude);


    peakFrequency_Hz = ...
        searchFrequency_Hz(localPeakIndex);
end


function addHarmonicLines( ...
    rotationFrequency_Hz, ...
    harmonicOrders)
% FFT 그래프에 1X, 2X, 3X 기준선을 표시합니다.

    if ~isfinite(rotationFrequency_Hz) || ...
            rotationFrequency_Hz <= 0

        return;
    end


    for iOrder = 1:numel(harmonicOrders)

        currentOrder = harmonicOrders(iOrder);


        harmonicFrequency_Hz = ...
            currentOrder * rotationFrequency_Hz;


        xline( ...
            harmonicFrequency_Hz, ...
            '--', ...
            sprintf('%dX', currentOrder), ...
            'LabelVerticalAlignment', 'bottom', ...
            'LabelHorizontalAlignment', 'left', ...
            'HandleVisibility', 'off');
    end
end


function sheetName = makeExcelSheetName(prefix, mainText)
% Excel 시트명에서 사용할 수 없는 문자를 제거하고,
% 최대 길이 31자 제한을 적용합니다.

    sheetName = [char(prefix), char(mainText)];


    invalidCharacters = { ...
        '\', ...
        '/', ...
        '*', ...
        '?', ...
        ':', ...
        '[', ...
        ']'};


    for iCharacter = 1:numel(invalidCharacters)

        sheetName = strrep( ...
            sheetName, ...
            invalidCharacters{iCharacter}, ...
            '_');
    end


    if length(sheetName) > 31

        sheetName = sheetName(1:31);
    end
end