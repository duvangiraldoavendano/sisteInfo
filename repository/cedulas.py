def cedulas_progra(programacion):
    archivo = open(programacion, 'r')
    lineas_espacios = archivo.readlines()
    archivo.close()

    lineas=[]
    for linea_espa in lineas_espacios:
        if(linea_espa!='\n'):
            lineas.append(linea_espa)

    cont=0
    lista=[]
    lineas_cedulas=[]
    for x in lineas:
        for char in x:
            if not(char in "1234567890"):
                cont=0
                lista=[]
            if char in "1234567890":
                cont+=1
                lista.append(char)
                if cont>5:
                    lineas_cedulas.append(x)
                    break

    lin_derecho=[]
    lin_cp=[]

    lin_DAD=[]
    lin_DDH=[]
    lin_DMD=[]
    lin_DPP=[]
    lin_DPR=[]
    lin_DRE=[]
    lin_DSS=[]
    lin_EDU=[]
    lin_PDD=[]

    def condicionales(linea_actual,indice_actual):
        cond=[]
        cond.append(False)
        cond.append(linea_actual[indice_actual].find(" DER ") == -1)
        cond.append(linea_actual[indice_actual].find(" DEP ") == -1)
        cond.append(linea_actual[indice_actual].find(" DEI ") == -1)
        cond.append(linea_actual[indice_actual].find(" CPT ") == -1)
        cond.append(linea_actual[indice_actual].find(" CPA ") == -1)
        cond.append(linea_actual[indice_actual].find(" DMD ") == -1)
        cond.append(linea_actual[indice_actual].find(" PDC ") == -1)
        cond.append(linea_actual[indice_actual].find(" DSS ") == -1)
        cond.append(linea_actual[indice_actual].find(" DRE ") == -1)
        cond.append(linea_actual[indice_actual].find(" DPP ") == -1)
        cond.append(linea_actual[indice_actual].find(" DAD ") == -1)
        cond.append(linea_actual[indice_actual].find(" DPR ") == -1)
        cond.append(linea_actual[indice_actual].find(" DDH ") == -1)
        cond.append(linea_actual[indice_actual].find(" PRA ") == -1)
        cond.append(linea_actual[indice_actual].find(" EDU ") == -1)
        cond.append(linea_actual[indice_actual].find(" PDD ") == -1)
        return cond

    index=0
    lineas=lineas_cedulas
    for linea in lineas:
        #Derecho
        condi1 = (linea.find(" DER ") != -1 and linea.find(" DER ") < 20)
        condi2 = (linea.find(" DEP ") != -1 and linea.find(" DEP ") < 20)
        condi3 = (linea.find(" DEI ") != -1 and linea.find(" DEI ") < 20)
        #Ciencias políticas
        condic1 = (linea.find(" CPT ") != -1) and (linea.find(" CPT ") < 20)
        condic2 = (linea.find(" CPA ") != -1) and (linea.find(" CPA ") < 20)
        #Posgrados
        condi_pos1=(linea.find(" DAD ") != -1) and (linea.find(" DAD ") < 20)#Especialización administrativo
        condi_pos2=(linea.find(" DDH ") != -1) and (linea.find(" DDH ") < 20)#Especialización derecho humanos
        condi_pos3=(linea.find(" DMD ") != -1) and (linea.find(" DMD ") < 20)#Maestria en Derecho
        condi_pos4=(linea.find(" DPP ") != -1) and (linea.find(" DPP ") < 20)#Especialización en derecho privado
        condi_pos5=(linea.find(" DPR ") != -1) and (linea.find(" DPR ") < 20)#Especialización en derecho procesal
        condi_pos6=(linea.find(" DRE ") != -1) and (linea.find(" DRE ") < 20)#Regionalización derecho
        condi_pos7=(linea.find(" DSS ") != -1) and (linea.find(" DSS ") < 20)#Especialización en Derecho seguridad social
        condi_pos8=(linea.find(" EDU ") != -1) and (linea.find(" EDU ") < 20)#Especialización en Derecho urbanistico
        condi_pos9=(linea.find(" PDD ") != -1) and (linea.find(" PDD ") < 20)#Doctorado en Derecho

        if(condi1 or condi2 or condi3):
            lin_derecho.append(linea)
            if(index<len(lineas)-1):
                index_int=index+1
            cond=condicionales(lineas,index_int)
            while(cond[1] and cond[2] and cond[3]  and cond[4] and cond[5] and cond[6] and cond[7] and cond[8] and cond[9] and cond[10] and cond[11] and cond[12] and cond[13] and cond[14] and cond[15] and cond[16]):
                lin_derecho.append(lineas[index_int])
                if (index_int == len(lineas)-1):
                    lin_derecho.append(lineas[index_int])
                    break
                index_int+=1
                cond = condicionales(lineas, index_int)
        elif (condic1 or condic2):
            lin_cp.append(linea)
            if(index<len(lineas)-1):
                index_int=index+1
            cond=condicionales(lineas,index_int)
            while (cond[1] and cond[2] and cond[3]  and cond[4] and cond[5] and cond[6] and cond[7] and cond[8] and cond[9] and cond[10] and cond[11] and cond[12] and cond[13] and cond[14] and cond[15] and cond[16]):
                lin_cp.append(lineas[index_int])
                if (index_int == len(lineas)-1):
                    lin_cp.append(lineas[index_int])
                    break
                index_int+=1
                cond = condicionales(lineas, index_int)
        elif(condi_pos1):
            lin_DAD.append(linea)
            if (index < len(lineas) - 1):
                index_int = index + 1
            cond=condicionales(lineas,index_int)
            while (cond[1] and cond[2] and cond[3]  and cond[4] and cond[5] and cond[6] and cond[7] and cond[8] and cond[9] and cond[10] and cond[11] and cond[12] and cond[13] and cond[14] and cond[15] and cond[16]):
                lin_DAD.append(lineas[index_int])
                if (index_int == len(lineas) - 1):
                    lin_DAD.append(lineas[index_int])
                    break
                index_int += 1
                cond=condicionales(lineas,index_int)
        elif (condi_pos2):
            lin_DDH.append(linea)
            if (index < len(lineas) - 1):
                index_int = index + 1
            cond=condicionales(lineas,index_int)
            while (cond[1] and cond[2] and cond[3]  and cond[4] and cond[5] and cond[6] and cond[7] and cond[8] and cond[9] and cond[10] and cond[11] and cond[12] and cond[13] and cond[14] and cond[15] and cond[16]):
                lin_DDH.append(lineas[index_int])
                if (index_int == len(lineas) - 1):
                    lin_DDH.append(lineas[index_int])
                    break
                index_int += 1
                cond = condicionales(lineas, index_int)
        elif (condi_pos3):
            lin_DMD.append(linea)
            if (index < len(lineas) - 1):
                index_int = index + 1
            cond=condicionales(lineas,index_int)
            while (cond[1] and cond[2] and cond[3]  and cond[4] and cond[5] and cond[6] and cond[7] and cond[8] and cond[9] and cond[10] and cond[11] and cond[12] and cond[13] and cond[14] and cond[15] and cond[16]):
                lin_DMD.append(lineas[index_int])
                if (index_int == len(lineas) - 1):
                    lin_DMD.append(lineas[index_int])
                    break
                index_int += 1
                cond = condicionales(lineas, index_int)
        elif (condi_pos4):
            lin_DPP.append(linea)
            if (index < len(lineas) - 1):
                index_int = index + 1
            cond = condicionales(lineas, index_int)
            while (cond[1] and cond[2] and cond[3]  and cond[4] and cond[5] and cond[6] and cond[7] and cond[8] and cond[9] and cond[10] and cond[11] and cond[12] and cond[13] and cond[14] and cond[15] and cond[16]):
                lin_DPP.append(lineas[index_int])
                if (index_int == len(lineas) - 1):
                    lin_DPP.append(lineas[index_int])
                    break
                index_int += 1
                cond = condicionales(lineas, index_int)
        elif (condi_pos5):
            lin_DPR.append(linea)
            if (index < len(lineas) - 1):
                index_int = index + 1
            cond = condicionales(lineas, index_int)
            while (cond[1] and cond[2] and cond[3]  and cond[4] and cond[5] and cond[6] and cond[7] and cond[8] and cond[9] and cond[10] and cond[11] and cond[12] and cond[13] and cond[14] and cond[15] and cond[16]):
                lin_DPR.append(lineas[index_int])
                if (index_int == len(lineas) - 1):
                    lin_DPR.append(lineas[index_int])
                    break
                index_int += 1
                cond = condicionales(lineas, index_int)
        elif (condi_pos6):
            lin_DRE.append(linea)
            if (index < len(lineas) - 1):
                index_int = index + 1
            cond = condicionales(lineas, index_int)
            while (cond[1] and cond[2] and cond[3]  and cond[4] and cond[5] and cond[6] and cond[7] and cond[8] and cond[9] and cond[10] and cond[11] and cond[12] and cond[13] and cond[14] and cond[15] and cond[16]):
                lin_DRE.append(lineas[index_int])
                if (index_int == len(lineas) - 1):
                    lin_DRE.append(lineas[index_int])
                    break
                index_int += 1
                cond = condicionales(lineas, index_int)
        elif (condi_pos7):
            lin_DSS.append(linea)
            if (index < len(lineas) - 1):
                index_int = index + 1
            cond = condicionales(lineas, index_int)
            while (cond[1] and cond[2] and cond[3]  and cond[4] and cond[5] and cond[6] and cond[7] and cond[8] and cond[9] and cond[10] and cond[11] and cond[12] and cond[13] and cond[14] and cond[15] and cond[16]):
                lin_DSS.append(lineas[index_int])
                if (index_int == len(lineas) - 1):
                    lin_DSS.append(lineas[index_int])
                    break
                index_int += 1
                cond = condicionales(lineas, index_int)
        elif (condi_pos8):
            lin_EDU.append(linea)
            if (index < len(lineas) - 1):
                index_int = index + 1
            cond = condicionales(lineas, index_int)
            while (cond[1] and cond[2] and cond[3]  and cond[4] and cond[5] and cond[6] and cond[7] and cond[8] and cond[9] and cond[10] and cond[11] and cond[12] and cond[13] and cond[14] and cond[15] and cond[16]):
                lin_EDU.append(lineas[index_int])
                if (index_int == len(lineas) - 1):
                    lin_EDU.append(lineas[index_int])
                    break
                index_int += 1
                cond = condicionales(lineas, index_int)
        elif (condi_pos9):
            lin_PDD.append(linea)
            if (index < len(lineas) - 1):
                index_int = index + 1
            cond = condicionales(lineas, index_int)
            while (cond[1] and cond[2] and cond[3]  and cond[4] and cond[5] and cond[6] and cond[7] and cond[8] and cond[9] and cond[10] and cond[11] and cond[12] and cond[13] and cond[14] and cond[15] and cond[16]):
                lin_PDD.append(lineas[index_int])
                if (index_int == len(lineas) - 1):
                    lin_PDD.append(lineas[index_int])
                    break
                index_int += 1
                cond = condicionales(lineas, index_int)
        else:
            pass
        index += 1

    def extrar_cedulas_de_lineas(linea_con_cedula):
        cedula_programa = []
        for lin_doc in linea_con_cedula:
            contador = 0
            cedula_indi_list = []
            for char in lin_doc:
                if not (char in "1234567890") and contador <= 5:
                    contador = 0
                    cedula_indi_list = []
                elif ((char in "1234567890" or char == " ")):
                    if (char == " " and contador > 5):
                        break
                    else:
                        cedula_indi_list.append(char)
                        contador += 1
                else:
                    pass
            cedula_indi_str = ''.join(cedula_indi_list)
            cedula_indi_int = int(cedula_indi_str)
            cedula_programa.append(cedula_indi_int)
        return cedula_programa

    cedula_dere = extrar_cedulas_de_lineas(lin_derecho)
    cedula_cp = extrar_cedulas_de_lineas(lin_cp)
    cedula_DAD = extrar_cedulas_de_lineas(lin_DAD)
    cedula_DDH = extrar_cedulas_de_lineas(lin_DDH)
    cedula_DMD = extrar_cedulas_de_lineas(lin_DMD)
    cedula_DPP = extrar_cedulas_de_lineas(lin_DPP)
    cedula_DPR = extrar_cedulas_de_lineas(lin_DPR)
    cedula_DRE = extrar_cedulas_de_lineas(lin_DRE)
    cedula_DSS = extrar_cedulas_de_lineas(lin_DSS)
    cedula_EDU = extrar_cedulas_de_lineas(lin_EDU)
    cedula_PDD = extrar_cedulas_de_lineas(lin_PDD)

    cedula_dere=set(cedula_dere)
    cedula_cp=set(cedula_cp)
    cedula_DAD=set(cedula_DAD)
    cedula_DDH = set(cedula_DDH)
    cedula_DMD = set(cedula_DMD)
    cedula_DPP = set(cedula_DPP)
    cedula_DPR = set(cedula_DPR)
    cedula_DRE = set(cedula_DRE)
    cedula_DSS = set(cedula_DSS)
    cedula_EDU = set(cedula_EDU)
    cedula_PDD = set(cedula_PDD)


    def convert_to_list(cedula_set):
        cedulasFinal = []
        for m in cedula_set:
            cedulasFinal.append(str(m))
        cedula_list = cedulasFinal
        return cedula_list

    cedula_dere = convert_to_list(cedula_dere)
    cedula_cp = convert_to_list(cedula_cp)
    cedula_DAD = convert_to_list(cedula_DAD)

    cedula_DAD = convert_to_list(cedula_DAD)
    cedula_DDH = convert_to_list(cedula_DDH)
    cedula_DPP = convert_to_list(cedula_DPP)
    cedula_DPR = convert_to_list(cedula_DPR)
    cedula_DRE = convert_to_list(cedula_DRE)
    cedula_DSS = convert_to_list(cedula_DSS)
    cedula_EDU = convert_to_list(cedula_EDU)
    cedula_PDD = convert_to_list(cedula_PDD)

    #return cedula_dere
    return cedula_cp,cedula_dere,lineas_cedulas,cedula_DAD