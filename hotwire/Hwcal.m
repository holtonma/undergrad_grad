% This program accepts two colums of data:
%
%				column 1 (x):  Flow Velocity    (in m/s)
%				colunm 2 (y):  Hot-Wire Voltage (in volts)
%
% The program takes this data, and ultimately produces a Polynomial Fit by
%  first fitting a third order polynomial of the form:
%				
%				U = Ao + A1E + A2E^2 + A3E^3    (i)
% 
% where U is the velocity data (originally taken from U-tube manometer measurements), and E
% represents Hot-Wire voltage values.
% 
% Once these original polynomial coefficients have been found, the values of the functions of
%
%				U^0.4, U^0.42, U^0.44, .....  U^0.6     (i.e. U^n)
% are found at each velocity in the calibration data.  Once these calibration points are known,
% they can be compared to their hot-wire voltage counterparts and an error analysis can be completed.
%
% The value of n that is found which yields the minimum error is then used to again fit the calibration
% data, and produce the final coefficients for the calibration polynomial above (i.)
%
% Written by:  Mark Holton             October 1999
%					holtonma@hotmail.com
%              Illinois Institute of Technology
%
%____________________________________________________________________________________________________________

% "FANCY stuff" to be done later:
% Prompt the user to input the filepath and name where the relevant data is located:
% fprintf (1,'Please indicate the filepath and filename of the Flow-velocity and Hot-Wire voltage value file to be calibrated']);
% check the number of data points that are present in the file
% Check number of data points--> assign to NumPts

DataFile = 'a:\hotwire\hwcltest';
NumPts   = 13   ;
NumChnls = 2    ;

% load DataFile
%calibrdata = zeros(NumChnls, NumPts);			   %Initializes data to all zeros
%fid	     = fopen([DataFile, '.dat'],'r');     %opens filename given by user and reads it
%calibrdata = fopen([DataFile, '.dat'],'r');		%opens file in DataFile variable and reads it
%fprintf (1,'Opened %s\n', [DataFile, '.dat']);  %Informs user that the filename has been opened
%calibrdata = fread(fid,[NumChnls, NumPts], 'float32');    %stores calibration data in file
%fclose(fid);
%fclose(calibrdata);
%fprintf(1,'Closed %s\n',[DataFile,'.dat']); 		%Informs user that the filename has been closed

load hwcltst.txt;
DataFile = hwcltst;

hwvolt = DataFile(:,2);
vel    = DataFile(:,1);
E = [ones(size(hwvolt)) hwvolt hwvolt.^2 hwvolt.^3]; 
												%creates matrix of colums of 1's, E's, E^2's, E^3's
A = E \ vel	;								%solves using Gaussian elim. for the coeff. matrix, A

% the following three lines plot a smooth curve from the first hw voltage pt. 
% to the last hw volt. pt. (taken at 0.05 intervals)
hw_curve = (hwvolt(1):0.0001:hwvolt(length(hwvolt)))';
plyline  = [ones(size(hw_curve)) hw_curve hw_curve.^2 hw_curve.^3]*A;
figure(1)
plot(hwvolt, vel,'o',hw_curve, plyline)
%figure(2)
%plot(hw_curve, plyline)

% Now, calculate the values for U^n: with n ranging from 0.40 to 0.60 (for comparison to
% (the hw voltage/constant side of the Polynomial Best Fit equation
vel40 = vel.^0.40;
vel41 = vel.^0.41;
vel42 = vel.^0.42;
vel43 = vel.^0.43;
vel44 = vel.^0.44;
vel45 = vel.^0.45;
vel46 = vel.^0.46;
vel47 = vel.^0.47;
vel48 = vel.^0.48;
vel49 = vel.^0.49;
vel50 = vel.^0.50;
vel51 = vel.^0.51;
vel52 = vel.^0.52;
vel53 = vel.^0.53;
vel54 = vel.^0.54;
vel55 = vel.^0.55;
vel56 = vel.^0.56;
vel57 = vel.^0.57;
vel58 = vel.^0.58;
vel59 = vel.^0.59;
vel60 = vel.^0.60;

% Now calculate the values each of the above vectors will be compared to:
% (i.e. calculate the right side of the equation, consisting of the E matrix and
% the A (constant) column vector [(4x4) * (4x1)]
hwALL = E*A;

% Now, compare vel40, vel41,...vel60 each to hwALL [(1x13)-(1x13)--> a (1x13) vector of
% the difference in values of the right side of the equation to the value of U^n at that pt.]
errvel40 = hwALL-vel40;
errvel41 = hwALL-vel41;
errvel42 = hwALL-vel42;
errvel43 = hwALL-vel43;
errvel44 = hwALL-vel44;
errvel45 = hwALL-vel45;
errvel46 = hwALL-vel46;
errvel47 = hwALL-vel47;
errvel48 = hwALL-vel48;
errvel49 = hwALL-vel49;
errvel50 = hwALL-vel50;
errvel51 = hwALL-vel51;
errvel52 = hwALL-vel52;
errvel53 = hwALL-vel53;
errvel54 = hwALL-vel54;
errvel55 = hwALL-vel55;
errvel56 = hwALL-vel56;
errvel57 = hwALL-vel57;
errvel58 = hwALL-vel58;
errvel59 = hwALL-vel59;
errvel60 = hwALL-vel60;

% Now, find the standard deviation of each of these error vectors
n40 = std(errvel40);
n41 = std(errvel41);
n42 = std(errvel42);
n43 = std(errvel43);
n44 = std(errvel44);
n45 = std(errvel45);
n46 = std(errvel46);
n47 = std(errvel47);
n48 = std(errvel48);
n49 = std(errvel49);
n50 = std(errvel50);
n51 = std(errvel51);
n52 = std(errvel52);
n53 = std(errvel53);
n54 = std(errvel54);
n55 = std(errvel55);
n56 = std(errvel56);
n57 = std(errvel57);
n58 = std(errvel58);
n59 = std(errvel59);
n60 = std(errvel60);

% Now, produce a column matrix of these std's
Nmatrix = [n40;n41;n42;n43;n44;n45;n46;n47;n48;n49;n50; ...
      n51;n52;n53;n54;n55;n56;n57;n58;n59;n60]
Nmatrix=Nmatrix'; %transposes matrix
minerror = min(Nmatrix)

Nvalue = 40:1:60;          %stores 40,41,42,43,....60 in a vector
figure(2)
plot(Nvalue,Nmatrix)
xlabel('value of n in U^n fit')
ylabel('std')
title('Exponent Determination for HW-Vel. Calibration')



                                    




