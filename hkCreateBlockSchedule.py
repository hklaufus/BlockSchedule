import xml.etree.ElementTree as ET
from datetime import date, timedelta
import svgwrite as SVG # pip install svgwrite
import sys

def CreateBlockSchedule(pFileName, pLevels = None):
    vProjectStart, vProjectFinish, vMilestones, vTasks = ReadMSPFile(pFileName, pLevels)
    vTasks = BuildTaskTree(vTasks)
    vTasks, vMaxWidth = BuildBlockSchedule(vProjectStart, vTasks)
    vMilestones = FilterMilestones(vMilestones)
    WriteSVG(pFileName, vProjectStart, vProjectFinish, vMaxWidth, vMilestones, vTasks)

def ReadMSPFile(pFileName, pLevels = None):
    vNameSpaces = "{http://schemas.microsoft.com/project}"
    vMyMilestones = []
    vMyTasks = []

    vTree = ET.parse(pFileName)
    vProject = vTree.getroot()

    try:
        vProjectStart = vProject.find(vNameSpaces+'StartDate').text.split('T')[0]
        vProjectStart2 = vProjectStart.split('-')
        vProjectStart3 = date(int(vProjectStart2[0]), int(vProjectStart2[1]), int(vProjectStart2[2]))
    except ValueError:
        print("Error in ReadMSPFile: vProjectStart invalid: ", vProjectStart)

    try:
        vProjectFinish = vProject.find(vNameSpaces+'FinishDate').text.split('T')[0]
        vProjectFinish2 = vProjectFinish.split('-')
        vProjectFinish3 = date(int(vProjectFinish2[0]), int(vProjectFinish2[1]), int(vProjectFinish2[2]))
    except ValueError:
        print("Error in ReadMSPFile: vTaskStart invalid: ", vProjectFinish)

    vTasks = vProject.find(vNameSpaces+'Tasks')
    for vTask in vTasks.iter(vNameSpaces+'Task'):
        vTaskOutlineLevel = int(vTask.find(vNameSpaces+'OutlineLevel').text)
        if(vTaskOutlineLevel>0):
            vTaskName = vTask.find(vNameSpaces+'Name').text
            vTaskOutlineNumber = vTask.find(vNameSpaces+'OutlineNumber').text

            try:
                vTaskStart = vTask.find(vNameSpaces+'Start').text.split('T')[0]
                vTaskStart2 = vTaskStart.split('-')
                vTaskStart3 = date(int(vTaskStart2[0]), int(vTaskStart2[1]), int(vTaskStart2[2]))
            except ValueError:
                print("Error in ReadMSPFile: vTaskStart invalid: ", vTaskStart)

            try:
                vTaskFinish = vTask.find(vNameSpaces+'Finish').text.split('T')[0]
                vTaskFinish2 = vTaskFinish.split('-')
                vTaskFinish3 = date(int(vTaskFinish2[0]), int(vTaskFinish2[1]), int(vTaskFinish2[2]))
            except ValueError:
                print("Error in ReadMSPFile: vTaskStart invalid: ", vTaskFinish)

            vTaskMilestone = vTask.find(vNameSpaces+'Milestone').text
            vTaskCritical = int(vTask.find(vNameSpaces+'Critical').text)

            # print(vTaskName, vTaskStart3, vTaskFinish3)

            if(vTaskMilestone == '1'):
                vMyMilestone = {}
                vMyMilestone['name'] = vTaskName
                vMyMilestone['startdate'] = vTaskStart3

                vMyMilestones.append(vMyMilestone)
            else:
                if((pLevels is None) or (vTaskOutlineLevel in pLevels)):
                    vMyTask = {}
                    vMyTask['name'] = vTaskName
                    vMyTask['startdate'] = vTaskStart3
                    vMyTask['finishdate'] = vTaskFinish3
                    vMyTask['outlinenumber'] = vTaskOutlineNumber
                    vMyTask['critical'] = vTaskCritical

                    vMyTasks.append(vMyTask)

    return vProjectStart3, vProjectFinish3, vMyMilestones, vMyTasks

def BuildTaskTree(pTasks):
    # Initialise parent pointer, number of children
    for vIdx, vTask in enumerate(pTasks):
        pTasks[vIdx]['parent'] = -1
        pTasks[vIdx]['children'] = []
        pTasks[vIdx]['numberdecendents'] = 0

    # Map children to parent, parent to children
    for vIdx1, vTask1 in enumerate(pTasks):
        vOutline1 = vTask1['outlinenumber'].split('.')

        for vIdx2, vTask2 in enumerate(pTasks):
            if (vIdx2 != vIdx1): # Skip if vTask2 == vTask1
                vOutline2 = vTask2['outlinenumber'].split('.')

                if(len(vOutline2) == len(vOutline1)+1) and (vOutline2[:-1] == vOutline1):
                    pTasks[vIdx2]['parent'] = vIdx1
                    pTasks[vIdx1]['children'].append(vIdx2)

    # Add total number of decendents
    vList1 = []
    for vIdx, vTask in enumerate(pTasks):
        # Add lowest level decendents to list, i.e. elements with no children.
        if(len(vTask['children'])==0):
            pTasks[vIdx]['numberdecendents'] = 0
            vList1.append(vIdx)

    # Loop through this list
    while(len(vList1)>0):
        vList2 = [] # temporary list
        for vIdx1 in vList1:
            # Parent of this element
            vIdx2 = pTasks[vIdx1]['parent']

            if(vIdx2!=-1):
                # Add numberdecendents of child to parent's numberdecendents
                pTasks[vIdx2]['numberdecendents'] = pTasks[vIdx2]['numberdecendents'] + pTasks[vIdx1]['numberdecendents'] + 1

                # Add parent to temprary list
                if(vIdx2 not in vList2):
                    vList2.append(vIdx2)

        # Reset vList1
        vList1 = vList2

    # Debug
    # for vTask in pTasks:
    #     print(vTask['name'], vTask['numberdecendents'], vTask['children'], vTask['parent'])

    return pTasks

def SetLeft(pTasks, pIdx, pOffset):
    # print(pTasks[pIdx]['name'])
    vCumulativeLeft = pTasks[pIdx]['left']

    for vIdx, vChildIdx in enumerate(pTasks[pIdx]['children']):
        pTasks[vChildIdx]['left'] = vCumulativeLeft + pOffset
        SetLeft(pTasks, vChildIdx, pOffset)
        vCumulativeLeft = vCumulativeLeft + pTasks[vChildIdx]['width'] + pOffset

def SetWidthSuccesors(pTasks, pIdx, pOffset):
    # print('Level: ' + str(pCurrentLevel) + '; name: ' + pTasks[pIdx]['name'] + '; pIdx = ' + str(pIdx))
    vNumberChildren = len(pTasks[pIdx]['children'])
    if(vNumberChildren>0):
        vChildWidth = (pTasks[pIdx]['width'] - (vNumberChildren+1)*pOffset) / vNumberChildren

        for vChildIdx in pTasks[pIdx]['children']:
            pTasks[vChildIdx]['width'] = vChildWidth
            SetWidthSuccesors(pTasks, vChildIdx, pOffset)

def UpdateWidthPredecessors(pTasks, pIdx, pDeltaWidth):
    # print(pTasks[pIdx]['name'], pTasks[pIdx]['width'])

    pTasks[pIdx]['width'] = pTasks[pIdx]['width'] + pDeltaWidth
    vParentIdx = pTasks[pIdx]['parent']
    if (vParentIdx >= 0):
        UpdateWidthPredecessors(pTasks, vParentIdx, pDeltaWidth)

def BuildBlockSchedule(pProjectStart, pTasks):
    vMinBlockWidth = 20
    vOffset = 5

    # Set top and height
    for vIdx, vTask in enumerate(pTasks):
        vLevel = len(vTask['outlinenumber'].split('.'))

        vTimeDifference1 = vTask['startdate'] - pProjectStart
        vTimeDifference2 = vTask['finishdate'] - vTask['startdate']

        pTasks[vIdx]['top'] = vTimeDifference1.days + vLevel*vOffset
        pTasks[vIdx]['height'] = (vTimeDifference2.days + 1) - vLevel*2*vOffset # NOTE: This can lead to height < 0 !!!

    # Search root where parent = -1
    vRoot = -1
    for vIdx, vTask in enumerate(pTasks):
        if(vTask['parent'] == -1):
            vRoot = vIdx

    if(vRoot==-1):
        sys.exit('ERROR in BuildBlockChart: Root not found..')

    # Loop through tree to set width values
    pTasks[vRoot]['width'] = 1000.0
    SetWidthSuccesors(pTasks, vRoot, vOffset)

    # Loop through tree to update width values based on minimum width
    for vIdx, vTask in enumerate(pTasks):
        vWidth = vTask['width']
        if (vWidth < vMinBlockWidth):
            vDeltaWidth = vMinBlockWidth - vWidth
            UpdateWidthPredecessors(pTasks, vIdx, vDeltaWidth)

    # Get the maximum width
    vMaxWidth = 0
    for vTask in pTasks:
        if(vTask['width']>vMaxWidth):
            vMaxWidth = vTask['width']

    # Set all left values
    pTasks[vRoot]['left'] = 0
    SetLeft(pTasks, vRoot, vOffset)

    return pTasks, vMaxWidth

def FilterMilestones(pMilestones):
    vDelete = []
    vIdx1 = 0
    while (vIdx1 < len(pMilestones)):
        vMilestone1 = pMilestones[vIdx1]
        vStartDate1 = vMilestone1['startdate']

        vIdx2 = vIdx1 + 1
        while (vIdx2 < len(pMilestones)):
            vMilestone2 = pMilestones[vIdx2]
            vStartDate2 = vMilestone2['startdate']

            if((vMilestone2 != vMilestone1) and (vStartDate2 == vStartDate1)):
                pMilestones[vIdx1]['name'] = pMilestones[vIdx1]['name'] + ' | ' + pMilestones[vIdx2]['name']
                vDelete.append(vMilestone2)

            vIdx2 = vIdx2 + 1

        vIdx1 = vIdx1 + 1

    for vMilestone in vDelete:
        if (vMilestone in pMilestones):
            pMilestones.remove(vMilestone)

    return pMilestones

def WriteSVG(pFileName, pProjectStart, pProjectFinish, pMaxWidth, pMilestones, pTasks):
    pFileName = pFileName + '.svg'
    vOffset = 5
    vMin = 0

    vDwg = SVG.Drawing(pFileName, profile='tiny')

    # Tasks
    vFillColours = [('#ffffff', '#cccccc'),('#ffaaaa', '#ffd5d5')]
    vStrokeColours = [('#000000', '#000000'),('#ff0000', '#ff0000')]

    for vTask in pTasks:
        # Define colours
        vFillColours2 = vFillColours[0]
        vStrokeColours2 = vStrokeColours[0]
        if((vTask['critical']) and (len(vTask['children']) == 0)):
            vFillColours2 = vFillColours[1]
            vStrokeColours2 = vStrokeColours[1]

        vFillColour = vFillColours2[0]
        vStrokeColour = vStrokeColours2[0]
        vLevel = len(vTask['outlinenumber'].split('.'))
        if(vLevel%2 == 0):
            # Even levels
            vFillColour = vFillColours2[1]
            vStrokeColour = vStrokeColours2[1]

        # Get task properties
        vName = vTask['name']
        vTop = vTask['top']
        vLeft = vTask['left']
        vHeight = vTask['height']
        if(vHeight<=0):
            vHeight=2

        vWidth = vTask['width']
        if(vWidth<=0):
            vWidth=2

        # Draw task
        vGroup = vDwg.g()
        if((vHeight>vMin) and (vWidth>vMin)):
            vRect = vDwg.rect(insert=(vLeft, vTop), size=(vWidth, vHeight), fill=vFillColour, stroke=vStrokeColour, stroke_width=2)
            vGroup.add(vRect)

            # Only draw text labels for lowest level tasks.
            if(len(vTask['children']) == 0):
                vFontSize = 12

                # Draw text horizontally or vertically
                vRotate = 0
                vTextLeft = vLeft + vWidth/2 -(len(vName)*vFontSize/2)/2
                vTextTop = vTop + vHeight - (vHeight-vFontSize)/2
                if(vHeight > vWidth):
                    vRotate = -90
                    vTextLeft = vLeft + vWidth - (vWidth-vFontSize)/2
                    vTextTop = vTop +vHeight/2 + (len(vName)*vFontSize/2)/2

                vText = vDwg.text(vName, insert=(vTextLeft, vTextTop), font_size=vFontSize, font_family='Arial')
                vText.rotate(vRotate, center=(vTextLeft, vTextTop))
                vGroup.add(vText)
        else:
            vCircle = vDwg.circle(center=(vLeft, vTop), r=vOffset, fill=vFillColour, stroke=vStrokeColour, stroke_width=2)
            vGroup.add(vCircle)

            vRotate = 0
            vTextLeft = vLeft + 2*vOffset
            vTextTop = vTop + vFontSize/2

            vText = vDwg.text(vName, insert=(vTextLeft, vTextTop), font_size=vFontSize, font_family='Arial')
            vText.rotate(vRotate, center=(vTextLeft, vTextTop))
            vGroup.add(vText)

        vDwg.add(vGroup)

    # Draw timeline
    vOffset = 20
    vStrokeColour = '#000000'
    vFillColour = '#000000'

    vStart = (-vOffset, 0)
    vEnd = (-vOffset, (pProjectFinish-pProjectStart).days)

    vLineV = vDwg.line(start=vStart, end=vEnd, fill=vFillColour, stroke=vStrokeColour, stroke_width=2)
    vDwg.add(vLineV)

    # Milestones
    vStrokeColour = '#000000'
    vFillColour = 'red'
    for vMilestone in pMilestones:
        vName = vMilestone['name']
        vStartDate = vMilestone['startdate']

        # Milestone symbol
        vDiamondLength = vOffset/2
        vCentre = (-vOffset, (vStartDate-pProjectStart).days - vDiamondLength*0.707)
        vDiamond = vDwg.rect(insert=vCentre, size=(vDiamondLength, vDiamondLength), fill=vFillColour, stroke=vStrokeColour, stroke_width=2)
        vDiamond.rotate(angle=45, center=vCentre)

        # Horizontal line
        vStart = (-vOffset + vDiamondLength*0.707, (vStartDate-pProjectStart).days)
        vEnd = (pMaxWidth + vOffset, (vStartDate-pProjectStart).days)
        vLineH = vDwg.line(start=vStart, end=vEnd, fill=vFillColour, stroke=vStrokeColour, stroke_width=1, stroke_dasharray=[8, 4, 2, 4], stroke_dashoffset=0)

        vString = vName + ' [' + vStartDate.strftime('%d-%m-%Y') + ']'
        vTextLeft = -1*(len(vString) * vFontSize/2 + 3*vOffset)
        vTextTop = (vStartDate-pProjectStart).days

        vText = vDwg.text(vString, insert=(vTextLeft, vTextTop))

        vGroup = vDwg.g()
        vGroup.add(vDiamond)
        vGroup.add(vLineH)
        vGroup.add(vText)
        vDwg.add(vGroup)


    print('Writing SVG to: ', pFileName)
    vDwg.save()
