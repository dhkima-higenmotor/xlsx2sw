import subprocess as sb
import win32com.client
import os
import pandas as pd
import pythoncom
import math
import numpy as np

class commSW:
    def __init__(self):
        self.drawing = self.drawing();
    #
    def startSW(self, *args):
        if not args:
            SW_PROCESS_NAME = r'C:/Program Files/SOLIDWORKS Corp/SOLIDWORKS/SLDWORKS.exe';
            sb.Popen(SW_PROCESS_NAME);
        else:
            year= int(args[0][-1]);
            SW_PROCESS_NAME = "SldWorks.Application.%d" % (20+(year-2));
            win32com.client.Dispatch(SW_PROCESS_NAME);
    #
    def shutSW(self):
        sb.call('Taskkill /IM SLDWORKS.exe /F');
    #
    def connectToSW(self):
        global swcom
        swcom = win32com.client.Dispatch("SLDWORKS.Application");
    #
    def openAssy(self, prtNameInp):                                                                  #
        self.prtName = prtNameInp;
        self.prtName = self.prtName.replace('\\','/');
        #
        if os.path.basename(self.prtName).split('.')[-1].lower() == 'sldasm': 
            pass;
        else:
            self.prtName+'.SLDASM'
        #
        openDoc     = swcom.OpenDoc6;
        arg1        = win32com.client.VARIANT(pythoncom.VT_BSTR, self.prtName);
        arg2        = win32com.client.VARIANT(pythoncom.VT_I4, 2);
        arg3        = win32com.client.VARIANT(pythoncom.VT_I4, 1);
        arg5        = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2);
        arg6        = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128);
        #
        openDoc(arg1, arg2, arg3, "", arg5, arg6);
    #
    def openPrt(self, prtNameInp):                                                                   #
        self.prtName = prtNameInp;
        self.prtName = self.prtName.replace('\\','/');
        #
        if os.path.basename(self.prtName).split('.')[-1].lower() == 'sldprt': 
            pass;
        else:
            self.prtName+'.SLDPRT'
        #
        openDoc     = swcom.OpenDoc6;
        arg1        = win32com.client.VARIANT(pythoncom.VT_BSTR, self.prtName);
        arg2        = win32com.client.VARIANT(pythoncom.VT_I4, 1);
        arg3        = win32com.client.VARIANT(pythoncom.VT_I4, 1);
        arg5        = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2);
        arg6        = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128);
        #
        openDoc(arg1, arg2, arg3, "", arg5, arg6);
    #
    def openDrw(self, drwNameInp):
        self.prtName = drwNameInp;
        self.prtName = self.prtName.replace('\\','/');
        #
        if os.path.basename(self.prtName).split('.')[-1].lower() == 'slddrw': 
            pass;
        else:
            self.prtName+'.SLDDRW'
        #
        openDoc     = swcom.OpenDoc6;
        arg1        = win32com.client.VARIANT(pythoncom.VT_BSTR, self.prtName);
        arg2        = win32com.client.VARIANT(pythoncom.VT_I4, 3);
        arg3        = win32com.client.VARIANT(pythoncom.VT_I4, 1);
        arg5        = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2);
        arg6        = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128);
        #
        openDoc(arg1, arg2, arg3, "", arg5, arg6);
    #
    def update(self):
        model   = swcom.ActiveDoc;
        model.EditRebuild3;
    #
    def closeDoc(self):
        swcom.CloseDoc(os.path.basename(self.prtName));
    #
    def save(self, directory, fileName, fileExtension):
        model   = swcom.ActiveDoc;
        directory   = directory.replace('\\','/');
        comFileName = directory+'/'+fileName+'.'+fileExtension;
        arg         = win32com.client.VARIANT(pythoncom.VT_BSTR, comFileName);
        model.SaveAs3(arg, 0, 0);
    #
    def getGlobalVars(self):
        model   = swcom.ActiveDoc;
        #
        eqMgr = model.GetEquationMgr;
        #
        n = eqMgr.getCount;
        #
        data = {};
        #
        for i in range(n):
            if eqMgr.GlobalVariable(i) == True:
                data[eqMgr.Equation(i).split('"')[1]] = i
            #
        #
        if len(data.keys()) == 0:
            raise KeyError("There are not any 'Global Variables' present in the currently active Solidworks document.");
        else:
            return data;
    #
    def modifyGlobalVar(self, variable, modifiedVal, unit):                                                         #
        model   = swcom.ActiveDoc;
        #
        eqMgr   = model.GetEquationMgr;
        #
        #data    = self.getGlobalVariables();
        data    = self.getGlobalVars();
        #
        if isinstance(variable, str) == True:
            eqMgr.Equation(data[variable], "\""+variable+"\" = "+str(modifiedVal)+unit+"");
        elif isinstance(variable, list) == True:
            if isinstance(modifiedVal, list) == True:
                if isinstance(unit, list) == True:
                    for i in range(len(variable)):
                        eqMgr.Equation(data[variable[i]], "\""+variable[i]+"\" = "+str(modifiedVal[i])+unit[i]+"");
                else:
                    raise TypeError("If a list of multiple variables is given, then lists of equal \n\
lengths should be given for 'modifiedVal' and 'unit' inputs.");
            else:
                raise TypeError("If a list of multiple variables is given, then lists of equal \n\
lengths should be given for 'modifiedVal' and 'unit' inputs.");
        else:
            raise TypeError("Incorrect input for the variables. Inputs can either be string, integer and string or lists containing variables, values and units.");
        #
        #self.updatePrt();
        self.update();
    #
    def modifyLinkedVar(self, variable, modifiedVal, unit, *args):
        if len(args) == 0:
            file = 'equations.txt';
        else:
            file = args[0];
        #
        # READ FILE WITH ORIGINAL DIMENSIONS
        try:
            reader      = open(file, 'r');
        except IOError:
            raise IOError;
        finally:
            data = {};
            numLines    = len(reader.readlines());
            reader.close();
            reader      = open(file);
            lines       = reader.readlines();
            reader.close();
            for i in range(numLines):
                dim     = lines[i].split('"')[1];
                tempVal = lines[i].split(' ')[1];
                #
                val     = tempVal.replace(unit,'').replace('= ','').replace('\n','');
                data[dim] = val;
        #
        # MODIFY DIMENSIONS
        if isinstance(variable, list) == True:
            if isinstance(modifiedVal, list) == True:
                if isinstance(unit, list) == True:
                    for z in range(len(variable)):
                        data[variable[i]] = modifiedVal[i];
                else:
                    raise TypeError("If a list of multiple variables is given, then lists of equal \n\
lengths should be given for 'modifiedVal' and 'unit' inputs.");
            else:
                raise TypeError("If a list of multiple variables is given, then lists of equal \n\
lengths should be given for 'modifiedVal' and 'unit' inputs.");
        elif isinstance(variable, str) == True:
            data[variable] = modifiedVal;
        else:
            raise TypeError("The inputs types given.");
        #
        # WRITE FILE WITH MODIFIED DIMENSIONS
        writer      = open(file, 'w');
        for key, value in data.items():
            writer.write('"'+key+'"= '+str(value)+unit);
        writer.close();
        #
        self.updatePrt();
    #
#
    class drawing:
        def __init__(self):
            pass;
        
        def _getTolType(self, tolTypeNum):
            tolType = {
                    1  : 'swTolBASIC',
                    2  : 'swTolBILAT',
                    10 : 'swTolBLOCK',
                    7  : 'swTolFIT',
                    9  : 'swTolFITTOLONLY',
                    8  : 'swTolFITWITHTOL',
                    11 : 'swTolGeneral',
                    3  : 'swTolLIMIT',
                    6  : 'swTolMAX',
                    7  : 'swTolMETRIC',
                    5  : 'swTolMIN',
                    0  : 'swTolNONE',
                    4  : 'swTolSYMMETRIC',
                    };
            
            return tolType[tolTypeNum];
        
        def _getDimType(self, dimTypeNum):
            dimType = {
                    3   : 'swAngularDimension',
                    4   : 'swArcLengthDimension',
                    10  : 'swChamferDimension',
                    6   : 'swDiameterDimension',
                    0   : 'swDimensionTypeUnknown',
                    11  : 'swHorLinearDimension',
                    7   : 'swHorOrdinateDimension',
                    2   : 'swLinearDimension',
                    1   : 'swOrdinateDimension',
                    5   : 'swRadialDimension',
                    13  : 'swScalarDimension',
                    12  : 'swVertLinearDimension',
                    8   : 'swVertOrdinateDimension',
                    9   : 'swZAxisDimension',
                    };
            
            if dimTypeNum == 0:
                return dimType[dimTypeNum].replace('sw','');
            else:
                return dimType[dimTypeNum].replace('sw','').replace('Dimension','');
        
        def _getDocUnits(self, swModel):
            dimUnit = {
                    'millimeters'   : 'mm',
                    'degrees'       : 'deg',
                    'meters'        : 'm',
                    'radians'       : 'rad',
                    };
            
            docUserUnitLinear   = swModel.GetUserUnit(win32com.client.VARIANT(pythoncom.VT_I4, 0));
            docUserUnitAngular  = swModel.GetUserUnit(win32com.client.VARIANT(pythoncom.VT_I4, 1));
            
            linearunit  = docUserUnitLinear.GetFullUnitName(True);
            angularunit = docUserUnitAngular.GetFullUnitName(True);
            
            return dimUnit[linearunit], dimUnit[angularunit];
        
        def getDimensions(self):
            swModel     = swcom.ActiveDoc;
            
            # ModelDocExtension = swModel.Extension;
            
            linearunit, angularunit = self._getDocUnits(swModel)
            
            swView      = swModel.GetFirstView;
            swView      = swView.GetNextView;

            data        = pd.DataFrame();
            dimName     = [];

            featureName = [];
            modelName   = [];
            dimValue    = [];
            dimType     = [];
            tolType     = [];
            maxTol      = [];
            minTol      = [];
            unit        = [];

            while isinstance(swView, win32com.client.CDispatch) == True:
                swDispDim   = swView.GetFirstDisplayDimension5;
                dimCount    = swView.GetDimensionCount4;
                
                #print(str(swView.GetDatumTargetSymCount)+'  '+str(swView.GetDisplayDimensionCount));
                #print(isinstance(swDispDim, win32com.client.CDispatch));
                
                for i in range(dimCount):
                                      
                    dimType.append(self._getDimType(swDispDim.Type2));
                                       
                    if self._getDimType(swDispDim.Type2) == 'Chamfer':
                        argAngle    = win32com.client.VARIANT(pythoncom.VT_I4, 0);
                        swDimAngle  = swDispDim.GetDimension2(argAngle);
                        argLength   = win32com.client.VARIANT(pythoncom.VT_I4, 1);
                        swDimLength = swDispDim.GetDimension2(argLength);
                        
                        dimName.append(swDimAngle.Name);
                        
                    else:
                        swDim = swDispDim.GetDimension;
                        dimName.append(swDim.Name);
                    
                    featureName.append(swDim.FullName.split('@')[1]);
                    modelName.append(swDim.FullName.split('@')[2]);

                    if self._getDimType(swDispDim.Type2) == 'Chamfer':
                        
                        dimValue.append(str(round(float(swDimLength.GetSystemValue2(""))*1e3,4))+' x '+str(round((((float(swDimAngle.GetSystemValue2("")))*(180))/math.pi)-90,4)));
                        unit.append(linearunit+'x'+angularunit);
                        tolType.append(self._getTolType(swDimAngle.GetToleranceType).replace('swTol','')+ ' x ' +self._getTolType(swDimLength.GetToleranceType).replace('swTol',''));
                        
                        if self._getTolType(swDimLength.GetToleranceType) == 'swTolSYMMETRIC':
                            
                            if self._getTolType(swDimAngle.GetToleranceType) == 'swTolSYMMETRIC':
                                maxTol.append(str(swDimAngle.Tolerance.GetMaxValue*1e3)+ ' x ' +str(round((((float(swDimLength.Tolerance.GetMaxValue))*(180))/math.pi)-0,4)));
                                minTol.append(str(-1*swDimAngle.Tolerance.GetMaxValue*1e3)+ ' x ' +str(-1*round((((float(swDimLength.Tolerance.GetMaxValue))*(180))/math.pi)-0,4)));
                            elif self._getTolType(swDimAngle.GetToleranceType) == 'swTolBILAT':
                                maxTol.append(str(swDimAngle.Tolerance.GetMaxValue*1e3)+ ' x ' +str(round((((float(swDimLength.Tolerance.GetMaxValue))*(180))/math.pi)-0,4)));
                                minTol.append(str(-1*swDimAngle.Tolerance.GetMaxValue*1e3)+ ' x ' +str(round((((float(swDimLength.Tolerance.GetMinValue))*(180))/math.pi)-0,4)));
                                
                        elif self._getTolType(swDim.GetToleranceType) == 'swTolBILAT':
                            
                            if self._getTolType(swDimAngle.GetToleranceType) == 'swTolSYMMETRIC':
                                maxTol.append(str(swDimAngle.Tolerance.GetMaxValue*1e3)+ ' x ' +str(round((((float(swDimLength.Tolerance.GetMaxValue))*(180))/math.pi)-0,4)));
                                minTol.append(str(swDimAngle.Tolerance.GetMinValue*1e3)+ ' x ' +str(-1*round((((float(swDimLength.Tolerance.GetMaxValue))*(180))/math.pi)-0,4)));
                            elif self._getTolType(swDimAngle.GetToleranceType) == 'swTolBILAT':
                                maxTol.append(str(swDimAngle.Tolerance.GetMaxValue*1e3)+ ' x ' +str(round((((float(swDimLength.Tolerance.GetMaxValue))*(180))/math.pi)-0,4)));
                                minTol.append(str(swDimAngle.Tolerance.GetMinValue*1e3)+ ' x ' +str(round((((float(swDimLength.Tolerance.GetMinValue))*(180))/math.pi)-0,4)));
                                
                        else:
                            maxTol.append(0);
                            minTol.append(0);
                        
                    elif self._getDimType(swDispDim.Type2) == 'Angular':
                        
                        dimValue.append(round((((float(swDim.GetSystemValue2("")))*(180))/math.pi),4));
                        unit.append(angularunit);
                        tolType.append(self._getTolType(swDim.GetToleranceType).replace('swTol',''));
                        
                        if self._getTolType(swDim.GetToleranceType) == 'swTolSYMMETRIC':
                            maxTol.append(swDim.Tolerance.GetMaxValue*1e3);
                            minTol.append(-1*swDim.Tolerance.GetMaxValue*1e3);
                        elif self._getTolType(swDim.GetToleranceType) == 'swTolBILAT':
                            maxTol.append(swDim.Tolerance.GetMaxValue*1e3);
                            minTol.append(swDim.Tolerance.GetMinValue*1e3);
                        elif self._getTolType(swDim.GetToleranceType) == 'swTolLIMIT':
                            maxTol.append(swDim.Tolerance.GetMaxValue*1e3);
                            minTol.append(swDim.Tolerance.GetMinValue*1e3);
                        else:
                            maxTol.append(0);
                            minTol.append(0);
                        
                    elif swDim.GetType == 0:
                        
                        dimValue.append(round(float(swDim.GetSystemValue2(""))*1e3,4));
                        unit.append(linearunit);
                        tolType.append(self._getTolType(swDim.GetToleranceType).replace('swTol',''));
                        
                        if self._getTolType(swDim.GetToleranceType) == 'swTolSYMMETRIC':
                            maxTol.append(swDim.Tolerance.GetMaxValue*1e3);
                            minTol.append(-1*swDim.Tolerance.GetMaxValue*1e3);
                        elif self._getTolType(swDim.GetToleranceType) == 'swTolBILAT':
                            maxTol.append(swDim.Tolerance.GetMaxValue*1e3);
                            minTol.append(swDim.Tolerance.GetMinValue*1e3);
                        elif self._getTolType(swDim.GetToleranceType) == 'swTolLIMIT':
                            maxTol.append(swDim.Tolerance.GetMaxValue*1e3);
                            minTol.append(swDim.Tolerance.GetMinValue*1e3);
                        else:
                            maxTol.append(0);
                            minTol.append(0);
                        
                    elif swDim.GetType == 2:
                        
                        dimValue.append(round(float(swDim.GetSystemValue2(""))*1e3,4));
                        unit.append(linearunit);
                        tolType.append(self._getTolType(swDim.GetToleranceType).replace('swTol',''));
                        
                        if self._getTolType(swDim.GetToleranceType) == 'swTolSYMMETRIC':
                            maxTol.append(swDim.Tolerance.GetMaxValue*1e3);
                            minTol.append(-1*swDim.Tolerance.GetMaxValue*1e3);
                        elif self._getTolType(swDim.GetToleranceType) == 'swTolBILAT':
                            maxTol.append(swDim.Tolerance.GetMaxValue*1e3);
                            minTol.append(swDim.Tolerance.GetMinValue*1e3);
                        elif self._getTolType(swDim.GetToleranceType) == 'swTolLIMIT':
                            maxTol.append(swDim.Tolerance.GetMaxValue*1e3);
                            minTol.append(swDim.Tolerance.GetMinValue*1e3);
                        else:
                            maxTol.append(0);
                            minTol.append(0);
                        
                    elif swDim.GetType == -1:
                        
                        dimValue.append(round(float(swDim.GetSystemValue2(""))*1e3,4));
                        unit.append(linearunit);
                        tolType.append(self._getTolType(swDim.GetToleranceType).replace('swTol',''));
                        
                        if self._getTolType(swDim.GetToleranceType) == 'swTolSYMMETRIC':
                            maxTol.append(swDim.Tolerance.GetMaxValue*1e3);
                            minTol.append(-1*swDim.Tolerance.GetMaxValue*1e3);
                        elif self._getTolType(swDim.GetToleranceType) == 'swTolBILAT':
                            maxTol.append(swDim.Tolerance.GetMaxValue*1e3);
                            minTol.append(swDim.Tolerance.GetMinValue*1e3);
                        elif self._getTolType(swDim.GetToleranceType) == 'swTolLIMIT':
                            maxTol.append(swDim.Tolerance.GetMaxValue*1e3);
                            minTol.append(swDim.Tolerance.GetMinValue*1e3);
                        else:
                            maxTol.append(0);
                            minTol.append(0);
                        
                    else:
                        pass;
                    #
                    swDispDim = swDispDim.GetNext3;

                swView = swView.GetNextView;
            
            data['DimensionName']   = np.asarray(dimName, dtype='<U64');
            data['FeatureName']     = np.asarray(featureName, dtype='<U64');
            data['ModelName']       = np.asarray(modelName, dtype='<U64');
            data['DimensionValue']  = np.asarray(dimValue, dtype='<U64');
            data['DimensionUnit']   = np.asarray(unit, dtype='<U64');
            data['DimensionType']   = np.asarray(dimType, dtype='<U64');
            data['ToleranceType']   = np.asarray(tolType, dtype='<U64');
            data['MaxTolerance']    = np.asarray(maxTol, dtype='<U64');
            data['MinTolerance']    = np.asarray(minTol, dtype='<U64');

            return data;
