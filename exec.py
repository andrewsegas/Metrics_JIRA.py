from protheus_process import CAProtheusProcess
import datetime
import numpy as np
import pandas as pd
import config
import xlwt
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

# ### Instância a classe de Auditoria de processo Protheus

ProtheusProcess = CAProtheusProcess()

#Format Date
styleDate = xlwt.XFStyle()
styleDate.num_format_str = 'DD-MM-YY hh:mm:ss'
#Format Number
styleNum = xlwt.XFStyle()
styleNum.num_format_str = '"0.00"'
#Format Negrito
styleTitle = xlwt.easyxf('font: name Arial, color-index red, bold on')


for x in range(0,len(config.projectList[0])):
    if config.projectList[0][x] in 'mob|ba':  # get story with manutencao
        cfilterCycle = '(' + config.projectList[1][x] + ') AND resolved >= startOfMonth(-1) AND resolved <= endOfMonth(-1) AND issuetype  in (Story, Manutenção) AND cf[17100] is not EMPTY ORDER BY resolved ASC'
        cfilterLead = '(' + config.projectList[1][x] + ') AND resolved >= startOfMonth(-1) AND resolved <= endOfMonth(-1) AND issuetype in (Story, Manutenção) AND cf[11043] is not EMPTY ORDER BY resolved ASC'
    else:
        #Cycle e Fila
        cfilterCycle = '(' + config.projectList[1][x] +') AND resolved >= startOfMonth(-1) AND resolved <= endOfMonth(-1) AND issuetype  in (Manutenção) AND cf[17100] is not EMPTY ORDER BY resolved ASC'
        #Lead e Suporte
        cfilterLead = '(' + config.projectList[1][x] +') AND resolved >= startOfMonth(-1) AND resolved <= endOfMonth(-1) AND issuetype in (Manutenção) AND cf[11043] is not EMPTY ORDER BY resolved ASC'

    # ### Obtém uma lista de issues conforme o filtro. ###
    print("Buscando issues no filtro para Cycle: " + config.projectList[0][x])
    issues = ProtheusProcess.getListIssues(cfilterCycle)
    print(f'Quantidade de Issues: {len(issues)}')

    issuesCycle = [ProtheusProcess.getCycle_Queue(x.key) for x in issues]

    d = []

    d.append(([x.key for x in issues]))  # 'Issue'
    d.append(([x[0] for x in issuesCycle]))  # 'Criação'
    d.append(([x[1] for x in issuesCycle]))  # 'Data Inicio Planejado'
    d.append(([x[2] for x in issuesCycle]))  # 'Resolvido'
    d.append(([x[3] for x in issuesCycle]))  # 'Cycle'
    d.append(([x[4] for x in issuesCycle]))  # 'Queue'

    wb = Workbook()
    sheet1 = wb.active
    sheet1.title = "Cycle"

    sheet1.cell(1, 1).value = "Issues"
    sheet1.cell(1, 1).font = Font(bold=True)
    sheet1.cell(1, 2).value = "Created"
    sheet1.cell(1, 2).font = Font(bold=True)
    sheet1.cell(1, 3).value = "Data Inicio Planejado"
    sheet1.cell(1, 3).font = Font(bold=True)
    sheet1.cell(1, 4).value = "Resolved"
    sheet1.cell(1, 4).font = Font(bold=True)
    sheet1.cell(1, 5).value = "Cycle"
    sheet1.cell(1, 5).font = Font(bold=True)
    sheet1.cell(1, 6).value = "Queue"
    sheet1.cell(1, 6).font = Font(bold=True)

    for i, l in enumerate(d):
        for j, col in enumerate(l):
            sheet1.cell(j+2, i+1, col)

    '''     if (i == 0):
                ws.write(j+1, i, col)
            elif (i < 4):
                ws.write(j+1, i, col.strftime("%d/%m/%Y %H:%M:%S"), styleDate)
            elif (i == 4):
                ws.write(j+1, i, xlwt.Formula(f"D{j+2}-C{j+2}"))
            else:
                ws.write(j+1, i, xlwt.Formula(f"C{j+2}-B{j+2}"))#str(col.days), styleNum)

    ws.write(j+2, 4, f"=MÉDIA(E2:E{j+2})")
    ws.write(j+2, 5, f"=MÉDIA(F2:F{j+2})")'''

    wb.save('C:/workspace/JIRA/metricas/Result/metrics_' + config.projectList[0][x] + '.xlsx')


    #Lead e Suporte

    # ### Obtém uma lista de issues conforme o filtro. ###
    print("Buscando issues no filtro para Lead: " + config.projectList[0][x])
    issues = ProtheusProcess.getListIssues(cfilterLead)
    print(f'Quantidade de Issues: {len(issues)}')

    issuesCycle = [ProtheusProcess.getLead_Suporte(x.key) for x in issues]

    ### Finished JIRA connection

    d = []

    d.append(([x.key for x in issues]))  # 'Issue'
    d.append(([x[0] for x in issuesCycle]))  # 'Criação'
    d.append(([x[1] for x in issuesCycle]))  # 'Data abertura ticket'
    d.append(([x[2] for x in issuesCycle]))  # 'Resolvido'
    d.append(([x[3] for x in issuesCycle]))  # 'Lead'
    d.append(([x[4] for x in issuesCycle]))  # 'Suporte'

    ws = wb.add_sheet('Lead')
    ws.write(0, 0, "Issues",styleTitle)
    ws.write(0, 1, "Created",styleTitle)
    ws.write(0, 2, "Data Abertura Ticket",styleTitle)
    ws.write(0, 3, "Resolved",styleTitle)
    ws.write(0, 4, "Lead",styleTitle)
    ws.write(0, 5, "Suporte",styleTitle)

    for i, l in enumerate(d):
        for j, col in enumerate(l):
            if (i == 0):
                ws.write(j+1, i, col)
            elif (i < 4):
                ws.write(j+1, i, col.strftime("%d/%m/%Y %H:%M:%S"), styleDate)
            elif (i == 4):
                ws.write(j+1, i, xlwt.Formula(f"D{j+2}-C{j+2}"))
            else:
                ws.write(j+1, i, xlwt.Formula(f"B{j+2}-C{j+2}"))#str(col.days), styleNum)

    ws.write(j+2, 4, f"=MÉDIA(E2:E{j+2})")
    ws.write(j+2, 5, f"=MÉDIA(F2:F{j+2})")


    wb.save('C:/workspace/JIRA/metricas/Result/metrics_'+ config.projectList[0][x] +'.xls')

ProtheusProcess.close()
