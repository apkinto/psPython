import ps
from ps import model
import datetime
import xml.etree.ElementTree as ET
import sys

def setVariables(config):
	variable = {}
	config = ET.parse(config)
	root = config.getroot()
	for var in root.find('variableList'):
		variable[var.tag] = var.text
	return variable

def getTime():
	currentTime = datetime.datetime.now()
	return currentTime

def openModel(inputDir, model):
	print "\nOPENING MODEL ", model, ":"
	start = getTime()
	psModel = inputDir + '/' + model
	app = ps.app.Application.instance()
	model = app.importModel(psModel)
	end = getTime()
	print "...Model opened:\t\t\t\t\t", end-start, "seconds"
	return model
	
def psSaveDxt(model, outputDir, modelName):
	name = modelName.split('.')
	dxtName = name[0] + ".dxt"
	#print dxtName, outputDir
	app = ps.app.Application.instance()
	app.saveModel(outputDir + "\\" + dxtName)
	
def	psSolve(model):
	print "SOLVING MODEL:"
	startSolve = getTime()
	model.solve()
	endSolve = getTime()
	solveTime = endSolve - startSolve
	print "...Solve Time:\t\t\t\t\t\t", solveTime, "seconds"

def	psGetAi(model, outputDir):
	print "Exporting Activity Instances:"
	start = getTime()
	outFile = outputDir + "/" + "ActivityInstances.txt"
	activityInstances = model.solution.getActivityInstances()			### Returns List instead of dictionary
	with open(outFile, "w") as f:
		for ai in activityInstances:
			#ai.getWorkOrderOperation()
			f.write("\t\t".join([ai.code, ai.display, ai.notes, str(ai.duration), str(ai.span), "\n"]))
	end = getTime()
	print "...Completed getting Acvity Instances:\t\t", end-start, "seconds\n"
	#v = ps.model.BusinessObjectVector()
	#v.extend(object for object in activityInstances)
	#print v

def	psGetWo(model, outputDir):
	print "Exporting Work Orders:"
	start = getTime()
	outFile = outputDir + "/" + "WO.txt"
	workOrders = model.getWorkOrders()				### returns dict {workOrderCode: woObjectt}
	with open(outFile, "w") as f:
		for w, wo in workOrders.iteritems():
			f.writelines([w, str(wo), "\n"])
 	end = getTime()
	print "...Completed getting WorkOrders:\t\t", end-start, "seconds\n"
	return workOrders

def	psGetWoAi(wo):
	woOp = {}
	woOp[wo.code] = wo.getPeggedInstances()
	#print woOp[wo]
	return woOp
	
def	psGetOp(model, outputDir):
	print "Exporting Operations:"
	start = getTime()
	outFile = outputDir + "/" + "Operations.txt"
	operations = model.getOperations()
	with open(outFile, "w") as f:
		for name, v in operations.iteritems():
			a = v.getActivity()
			print a.code
			attrValues = ps.model.Activity.getAttributeValues(a)
			for av in attrValues:
				f.writelines([name, 2*"\t", v.code, 2*"\t", av.code, "\n"])
				print name, v.code, "--->", av.code
			#v.getAttributeValues()
	end = getTime()
	print "...Completed getting Operations:\t\t", end-start, "seconds\n"

def	psGetAttr(model, outputDir):
	print "Exporting Attributes:"
	start = getTime()
	outFile = outputDir + "/" + "Attributes.txt"
	attr = model.getAttributes()
	with open(outFile, "w") as f:
		for a, v in attr.iteritems():
			f.writelines([a, str(v), "\n"])
			#print "\n\n!!!", a, v, "\n\n"
	end = getTime()
	print "...Completed getting Attributes:\t\t", end-start, "seconds\n"
	#val = ps.model.Attribute.getValues(attr['Colour'])
	#print val[0].code, val[0].display, val[0].notes

def psRemoveIdle(model, res):
	resource= model.findResource(res)
	resSeq = model.solution.findResourceSchedule(resource)
	model.solution.repairMode = ps.ms.repairMode.unconstrained		# Set to Whiteboard mode
	manualSchedule = ps.ms.service(model.schedule)
	manualSchedule.pasteActivityInstances([resSeq[1], resSeq[0]], resource, datetime.datetime(2014, 4, 14, 10, 0), ps.ms.resequencingMode.singleStage)
	model.repair()	
	resSeq = model.solution.findResourceSchedule(resource)
	for ai in resSeq:
		print ai.code, ai.display, ai.span
		
	
'''
def	psManSched(model, woNumber, operation):
	woAi = psGetWoAi(WO[woNumber])
	i = 0
	for ai in woAi[woNumber]:
		if ai.code == operation:
			if i == 0:
				#start = ai.span[1] + datetime.timedelta(0,1)
				ai.span[0] = ai.span[0] + datetime.timedelta(0,1)
				ai.span[1] = ai.span[1] + datetime.timedelta(0,1)
			else:
				print start, start + ai.duration
				ai.span[0] = start
				ai.span[1] = start + ai.duration
			#x = ai.duration + ai.duration # datetime.timedelta(0,1)
				print i, ai.code, ai.span[0], ai.span[1], ai.duration, "--", start
			#start = start + ai.duration + datetime.timedelta(0,1)
			i = i + 1		#print start
	#print woAi['1']
'''   
###################			Main				#########################

if __name__ == "__main__":

	variables = setVariables('psPythonConfig.xml')
	outDir = variables['psOutputDirectory']
	modelXml = variables['psXml']
	myModel = openModel(variables['psInputDirectory'], modelXml)
	psSolve(myModel) 
	psGetAi(myModel, outDir)
	psGetAttr(myModel, outDir)
	WO = psGetWo(myModel, outDir)
	#a = ps.model.AttributeVector()
	psGetOp(myModel, outDir)
	psRemoveIdle(myModel, variables['resource'])
	#psManSched(myModel, '1', 'op20')
	
	psSaveDxt(myModel, outDir, modelXml)
	

'''	
	getWoActivities(myModel, variables['psOutputDirectory'])
	resource = sys.argv[2]
	resource = variables['resource']
		getResourceSchedule(myModel, variables['psOutputDirectory'], resource)
'''
