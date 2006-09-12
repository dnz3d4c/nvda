import struct
import re
import win32api
import win32com
import win32con
import win32gui
import pyAA
import debug
import audio
from keyEventHandler import key, sendKey
from constants import *
from config import conf
import dictionaries
import api


#Some api functions specific to NVDAObjects

def getRoleName(role):
	if dictionaries.roleNames.has_key(role) is True:
		return dictionaries.roleNames[role]
	else:
		return role

def createStateList(stateBits):
	stateList=[]
	for bitPos in range(32):
		bitVal=1<<bitPos
		if stateBits&bitVal:
			stateList+=[bitVal]
	return stateList

def getStateNames(states,opposite=False):
	str=""
	for state in createStateList(states):
		str="%s %s"%(str,getStateName(state,opposite=opposite))
	return str

def getStateName(state,opposite=False):
	if dictionaries.stateNames.has_key(state):
		name=dictionaries.stateNames[state]
	else:
		name=state
	if opposite is True:
		name="not %s"%name
	return name

#The classes

class NVDAObject(object):

	def __new__(cls,*args):
		if (len(args)!=1) or not isinstance(args[0],pyAA.AA.AccessibleObject):
			debug.writeError("class takes an object of type pyAA.AA.AccessibleObject as its only parameter")
			return None
		accObject=args[0]
		try:
			window=accObject.Window
		except:
			return None
		if not win32gui.IsWindow(window):
			return None
		className=win32gui.GetClassName(window)
		try:
			role=accObject.GetRole()
		except:
			return None
		NVDAClass=classMap.get("%s_%s"%(className,role),None)
		if not NVDAClass:
			NVDAClass=classMap.get("%s"%className,None)
			if not NVDAClass:
				NVDAClass=cls
		return object.__new__(NVDAClass,*args)

	def __init__(self,accObject):
		self.accObject=accObject
		self.keyMap={}
		self.lastStates=self.getStates()

	def __eq__(self,other):
		if (self.getProcessID()==other.getProcessID()) and (self.getWindowHandle()==other.getWindowHandle()) and (self.getRole()==other.getRole()) and (self.getChildID()==other.getChildID()) and (self.getLocation()==other.getLocation()):
			return True
		else:
			return False

	def __ne__(self,other):
		if (self.getProcessID()!=other.getProcessID()) or (self.getWindowHandle()!=other.getWindowHandle()) or (self.getRole()!=other.getRole()) or (self.getChildID()!=other.getChildID()) or (self.getLocation()!=other.getLocation()):
			return True
		else:
			return False

	def speakObject(obj):
		window=obj.getWindowHandle()
		name=obj.getName()
		typeString=obj.getTypeString()
		stateNames=getStateNames(obj.filterStates(obj.getStates()))
		value=obj.getValue()
		description=obj.getDescription()
		if description==name:
			description=None
		help=obj.getHelp()
		if conf["presentation"]["reportKeyboardShortcuts"]:
			keyboardShortcut=obj.getKeyboardShortcut()
		else:
			keyboardShortcut=None
		position=None
		childID=obj.getChildID()
		if childID>0:
			parent=obj.getParent()
			if parent:
				parentChildCount=parent.getChildCount()
				position="%s of %s"%(childID,parentChildCount)
		#if role!=ROLE_SYSTEM_GROUPING:
		#	groupName=getObjectGroupName(accObject)
		#else:
		#	groupName=None
		groupName=None
		audio.speakObjectProperties(groupName=groupName,name=name,typeString=typeString,stateText=stateNames,value=value,description=description,help=help,keyboardShortcut=keyboardShortcut,position=position)

	def getWindowHandle(self):
		try:
			window=self.accObject.Window
		except:
			return None
		return window

	def getName(self):
		try:
			name=self.accObject.GetName()
		except:
			name=""
		if not name:
			window=self.getWindowHandle()
			if not window:
				return ""
			name=win32gui.GetWindowText(window)
		return name

	def getValue(self):
		value=None
		try:
			value=self.accObject.GetValue()
		except:
			pass
		if value:
			return value
		else:
			return ""

	def getRole(self):
		try:
			return self.accObject.GetRole()
		except:
			return ""

	def getTypeString(self):
		role=self.getRole()
		if conf["presentation"]["reportClassOfAllObjects"] or (conf["presentation"]["reportClassOfClientObjects"] and (role==ROLE_SYSTEM_CLIENT)):
			typeString=self.getClassName()
		else:
			typeString=""
		return typeString+" %s"%getRoleName(self.getRole())


	def getStates(self):
		states=0
		try:
			states=self.accObject.GetState()
		except:
			pass
		return states

	def filterStates(self,states):
		states-=(states&STATE_SYSTEM_FOCUSED)
		states-=(states&STATE_SYSTEM_FOCUSABLE)
		states-=(states&STATE_SYSTEM_SELECTABLE)
		states-=(states&STATE_SYSTEM_MULTISELECTABLE)
		states-=(states&STATE_SYSTEM_READONLY)
		states-=(states&STATE_SYSTEM_INVISIBLE)
		states-=(states&STATE_SYSTEM_HOTTRACKED)
		states-=(states&STATE_SYSTEM_OFFSCREEN)
		states-=(states&STATE_SYSTEM_DEFAULT)
		return states

	def getDescription(self):
		try:
			return self.accObject.GetDescription()
		except:
			return ""

	def getHelp(self):
		try:
			return self.accObject.GetHelp()
		except:
			return ""

	def getKeyboardShortcut(self):
		keyboardShortcut=None
		try:
			keyboardShortcut=self.accObject.GetKeyboardShortcut()
		except:
			return ""
		if not keyboardShortcut:
			return ""
		else:
			return keyboardShortcut

	def getChildID(self):
		try:
			return self.accObject.child
		except:
			return None

	def getChildCount(self):
		return len(self.getChildren())

	def getProcessID(self):
		try:
			return self.accObject.ProcessID
		except:
			debug.writeException("NVDAObjects.NVDAObject.getProcessID")
			return None

	def getLocation(self):
		try:
			return self.accObject.GetLocation()
		except:
			return None

	def getClassName(self):
		return win32gui.GetClassName(self.getWindowHandle())

	def getParent(self):
		try:
			accObject=self.accObject.GetParent()
		except:
			return None
		if accObject.GetRole()==pyAA.Constants.ROLE_SYSTEM_WINDOW:
			try:
				return NVDAObject(accObject.GetParent())
			except:
				return None
		else:
			return NVDAObject(accObject)

	def getNext(self,checkReverse=True):
		try:
			parentObject=NVDAObject(self.accObject.GetParent())
			parentRole=parentObject.getRole()
		except:
			parentObject=None
			parentRole=None
		if parentObject and (parentRole==ROLE_SYSTEM_WINDOW):
			try:
				nextObject=NVDAObject(parentObject.accObject.Navigate(pyAA.Constants.NAVDIR_NEXT))
			except:
				debug.writeError("NVDAObject.getNext: failed to get next window object")
				return None
			if checkReverse:
				try:
					testObject=NVDAObject(nextObject.accObject.Navigate(pyAA.Constants.NAVDIR_PREVIOUS))
				except:
					debug.writeError("NVDAObject.getNext: failed to get test object from next window object")
					return None
			else:
				testObject=None
			if nextObject and (nextObject!=parentObject) and (not checkReverse or (testObject and (testObject==parentObject))):  
				return NVDAObject(pyAA.AccessibleObjectFromWindow(nextObject.getWindowHandle(),pyAA.Constants.OBJID_CLIENT))
		else:
			try:
				nextObject=NVDAObject(self.accObject.Navigate(pyAA.Constants.NAVDIR_NEXT))
				if checkReverse:
					testObject=NVDAObject(nextObject.accObject.Navigate(pyAA.Constants.NAVDIR_PREVIOUS))
				else:
					testObject=None
				if nextObject and (nextObject!=self) and (not checkReverse or (testObject and (testObject==self))):
					return nextObject
				else:
					return None
			except:
				return None

	def getPrevious(self):
		try:
			parentObject=NVDAObject(self.accObject.GetParent())
			parentRole=parentObject.getRole()
		except:
			parentObject=None
			parentRole=None
		if parentObject and (parentRole==ROLE_SYSTEM_WINDOW):
			try:
				prevObject=NVDAObject(parentObject.accObject.Navigate(pyAA.Constants.NAVDIR_PREVIOUS))
			except:
				debug.writeError("NVDAObject.getPrevious: failed to get previous window object")
				return None
			try:
				testObject=NVDAObject(prevObject.accObject.Navigate(pyAA.Constants.NAVDIR_NEXT))
			except:
				debug.writeError("NVDAObject.getPrevious: failed to get test object from previous window object")
				return None
			if prevObject and testObject and (testObject==parentObject) and (prevObject!=parentObject):  
				return NVDAObject(pyAA.AccessibleObjectFromWindow(prevObject.getWindowHandle(),pyAA.Constants.OBJID_CLIENT))
		else:
			try:
				prevObject=NVDAObject(self.accObject.Navigate(pyAA.Constants.NAVDIR_PREVIOUS))
				testObject=NVDAObject(prevObject.accObject.Navigate(pyAA.Constants.NAVDIR_NEXT))
				if prevObject and testObject and (self==testObject) and (prevObject!=self):
					return prevObject
				else:
					return None
			except:
				return None

	def getFirstChild(self):
		try:
			childObject=self.accObject.Navigate(pyAA.Constants.NAVDIR_FIRSTCHILD)
			if childObject.GetRole()==pyAA.Constants.ROLE_SYSTEM_WINDOW:
				childObject=pyAA.AccessibleObjectFromWindow(childObject.Window,pyAA.Constants.OBJID_CLIENT)
			testObject=self.accObject
			if childObject and ((childObject.Window!=testObject.Window) or (childObject.GetRole()!=testObject.GetRole()) or (childObject.ChildID!=testObject.ChildID)):
				return NVDAObject(childObject)
			else:
				return None
		except:
			return None

	def doDefaultAction(self):
		try:
			self.accObject.DoDefaultAction()
		except:
			pass

	def getChildren(self):
		children=[]
		obj=self.getFirstChild()
		if obj:
			children.append(obj)
			next=obj.getNext()
			while next:
				children.append(next)
				next=next.getNext()
		return children

	def getActiveChild(self):
		try:
			child=self.accObject.GetFocus()
		except:
			return None
		return NVDAObject(child)

	def hasFocus(self):
		states=0
		try:
			states=self.accObject.GetState()
		except:
			pass
		if states&pyAA.Constants.STATE_SYSTEM_FOCUSED:
			return True
		else:
			return False

	def getCharacter(self,index=None):
		if index is None:
			index=self.getCaretIndex()
		if index[1]>=self.getLineLength(index=index):
			return None
		return self.getLine(index=index)[index[1]]

	def getWord(self,index=None):
		if not index:
			index=self.getCaretIndex()
		end=self.getWordEndIndex(index)
		if not end or (end==index):
			text=self.getCharacter(index=index)
		else:
			text=self.getTextRange(index,end)
		return text

	def event_foreground(self):
		api.setVirtualBuffer(self.getWindowHandle())
		api.setVirtualBufferCursor(api.getVirtualBuffer().getCaretIndex())
		self.speakObject()

	def event_focusObject(self):
		if api.getVirtualBuffer().getWindowHandle()!=api.getForegroundWindow():
			api.setVirtualBuffer(api.getForegroundWindow())
		api.setVirtualBufferCursor(api.getVirtualBuffer().getCaretIndex())
		if self.hasFocus():
			self.speakObject()

	def event_objectValueChange(self):
		audio.speakObjectProperties(value=self.getValue())

	def event_objectStateChange(self):
		states=self.getStates()
		if states is None:
			return None
		states_on=states-(states&self.lastStates)
		audio.speakObjectProperties(stateText=getStateNames(self.filterStates(states_on)))
		states_off=self.lastStates-(states&self.lastStates)
		audio.speakObjectProperties(stateText=getStateNames(self.filterStates(states_off),opposite=True))
		self.lastStates=states

class NVDAObject_dialog(NVDAObject):

	def event_foreground(self):
		self.speakObject()
		for child in self.getChildren():
			states=child.getStates()
			if (not states&STATE_SYSTEM_OFFSCREEN) and (not states&STATE_SYSTEM_INVISIBLE) and (not states&STATE_SYSTEM_UNAVAILABLE):
				child.speakObject()
			if child.getRole()==ROLE_SYSTEM_PROPERTYPAGE:
				for grandChild in child.getChildren():
					states=grandChild.getStates()
					if (not states&STATE_SYSTEM_OFFSCREEN) and (not states&STATE_SYSTEM_INVISIBLE) and (not states&STATE_SYSTEM_UNAVAILABLE):
						grandChild.speakObject()


class NVDAObject_Shell_TrayWnd(NVDAObject):

	def event_foreground(self):
		pass

class NVDAObject_Progman(NVDAObject):

	def event_foreground(self):
		pass

class NVDAObject_Edit(NVDAObject):

	def __init__(self,accObject):
		NVDAObject.__init__(self,accObject)
		self.keyMap={
			key("ExtendedUp"):self.script_moveByLine,
			key("ExtendedDown"):self.script_moveByLine,
			key("ExtendedLeft"):self.script_moveByCharacter,
			key("ExtendedRight"):self.script_moveByCharacter,
			key("Control+ExtendedLeft"):self.script_moveByWord,
			key("Control+ExtendedRight"):self.script_moveByWord,
			key("Shift+ExtendedRight"):self.script_changeSelection,
			key("Shift+ExtendedLeft"):self.script_changeSelection,
			key("Shift+ExtendedHome"):self.script_changeSelection,
			key("Shift+ExtendedEnd"):self.script_changeSelection,
			key("Shift+ExtendedUp"):self.script_changeSelection,
			key("Shift+ExtendedDown"):self.script_changeSelection,
			key("Control+Shift+ExtendedLeft"):self.script_changeSelection,
			key("Control+Shift+ExtendedRight"):self.script_changeSelection,
			key("ExtendedHome"):self.script_moveByCharacter,
			key("ExtendedEnd"):self.script_moveByCharacter,
			key("ExtendedDelete"):self.script_delete,
			key("Back"):self.script_backspace,
		}

	def getCaretIndecies(self):
		word=win32gui.SendMessage(self.getWindowHandle(),win32con.EM_GETSEL,0,0)
		if word<0:
			debug.writeError("window.getCaretIndex: got invalid selection word from window")
			return None
		curPos=win32api.LOWORD(word)
		lineNum=win32gui.SendMessage(self.getWindowHandle(),win32con.EM_LINEFROMCHAR,curPos,0)
		linePos=win32gui.SendMessage(self.getWindowHandle(),win32con.EM_LINEINDEX,lineNum,0)
		startIndex=(lineNum,curPos-linePos)
		curPos=win32api.HIWORD(word)
		lineNum=win32gui.SendMessage(self.getWindowHandle(),win32con.EM_LINEFROMCHAR,curPos,0)
		linePos=win32gui.SendMessage(self.getWindowHandle(),win32con.EM_LINEINDEX,lineNum,0)
		endIndex=(lineNum,curPos-linePos)
		point=(startIndex,endIndex)
		return point

	def getCaretIndex(self):
		point=self.getCaretIndecies()
		return point[1]

	def getLineCount(self):
		lineCount=win32gui.SendMessage(self.getWindowHandle(),win32con.EM_GETLINECOUNT,0,0)
		if lineCount<0:
			debug.writeError("window.getLineCount: failed to get line count")
			return None
		return lineCount


	def getLineLength(self,index=None):
		if index is None:
			index=self.getCaretIndex()
		lineLength=win32gui.SendMessage(self.getWindowHandle(),win32con.EM_LINELENGTH,win32gui.SendMessage(self.getWindowHandle(),win32con.EM_LINEINDEX,index[0],0),0)
		if lineLength<0:
			debug.writeError("window.getLineLength: line length invalid or negative (line number %d, line position %d"%(lineNum,curPos))
			return None
		return lineLength

	def getLine(self,index=None):
		if index is None:
			index=self.getCaretIndex()
		lineNum=index[0]
		lineLength=self.getLineLength(index=index)
		if lineLength is None:
			debug.writeError("window.getLine: line length is not valid")
			return None
		if lineLength==0:
			return None
		lineBuf=struct.pack('i',lineLength+1)
		lineBuf=lineBuf+"".ljust(lineLength-2)
		res=win32gui.SendMessage(self.getWindowHandle(),win32con.EM_GETLINE,lineNum,lineBuf)
		line="%s"%lineBuf[0:lineLength]
		return line

	def getSelection(self,indecies=None):
		if indecies is None:
			indecies=self.getCaretIndecies()
		if indecies is None:
			debug.writeError("window.getCharacter: failed to get index")
			return None
		if indecies[0]==indecies[1]:
			debug.writeError("window.getSelection: no selection")
			return None
		selection=self.getTextRange(indecies[0],indecies[1])
		if selection is None:
			return None
		return selection

	def script_moveByLine(self,keyPress):
		sendKey(keyPress)
		audio.speakText(self.getLine())

	def script_moveByCharacter(self,keyPress):
		sendKey(keyPress)
		audio.speakSymbol(self.getCharacter())

	def script_moveByWord(self,keyPress):
		sendKey(keyPress)
		audio.speakText(self.getWord())

	def script_changeSelection(self,keyPress):
		selectionPoints=self.getCaretIndecies()
		if selectionPoints[0]==selectionPoints[1]:
			selectionPoints=None
		sendKey(keyPress)
		newSelectionPoints=self.getCaretIndecies()
		if newSelectionPoints[0]==newSelectionPoints[1]:
			newSelectionPoints=None
		if newSelectionPoints and not selectionPoints:
			audio.speakText("selected %s"%self.getTextRange(newSelectionPoints[0],newSelectionPoints[1]))
		elif not newSelectionPoints:
			audio.speakSymbol(self.getCharacter())
		elif selectionPoints and newSelectionPoints: 
			if newSelectionPoints[1]>selectionPoints[1]:
				audio.speakText("selected %s"%self.getTextRange(selectionPoints[1],newSelectionPoints[1]))
			elif newSelectionPoints[0]>selectionPoints[0]:
				audio.speakText("unselected %s"%self.getTextRange(selectionPoints[0],newSelectionPoints[0]))
			elif newSelectionPoints[1]<selectionPoints[1]:
				audio.speakText("unselected %s"%self.getTextRange(newSelectionPoints[1],selectionPoints[1]))
			elif newSelectionPoints[0]<selectionPoints[0]:
				audio.speakText("selected %s"%self.getTextRange(newSelectionPoints[0],selectionPoints[0]))

	def script_delete(self,keyPress):
		sendKey(keyPress)
		audio.speakSymbol(self.getCharacter())

	def script_backspace(self,keyPress):
		point=self.getCaretIndex()
		if not point==[0,0]: 
			delChar=self.getCharacter(index=self.getPreviousCharacterIndex(point))
			sendKey(keyPress)
			newPoint=self.getCaretIndex()
			if newPoint<point:
				audio.speakSymbol(delChar)
		else:
			sendKey(keyPress)

	def event_objectValueChange(self):
		pass

class NVDAObject_checkBox(NVDAObject):

	def getStates(self):
		states=NVDAObject.getStates(self)
		states-=states&pyAA.Constants.STATE_SYSTEM_PRESSED
		return states

class NVDAObject_mozillaUIWindowClass(NVDAObject):

	def event_focusObject(self):
		if api.getVirtualBuffer().getWindowHandle()!=api.getForegroundWindow():
			api.setVirtualBuffer(api.getForegroundWindow())
		api.setVirtualBufferCursor(api.getVirtualBuffer().getCaretIndex())
		self.speakObject()

class NVDAObject_mozillaUIWindowClass_application(NVDAObject_mozillaUIWindowClass):

	def getValue(self):
		return ""

	def event_focusObject(self):
		if api.getVirtualBuffer().getWindowHandle()!=api.getForegroundWindow():
			api.setVirtualBuffer(api.getForegroundWindow())
		api.setVirtualBufferCursor(api.getVirtualBuffer().getCaretIndex())

	def getFirstChild(self):
		try:
			children=self.accObject.GetChildren()
		except:
			return None
		for child in children:
			try:
				role=child.GetRole()
				if role not in [ROLE_SYSTEM_TOOLTIP,ROLE_SYSTEM_MENUPOPUP]:
					return NVDAObject(child)
			except:
				pass


class NVDAObject_mozillaContentWindowClass(NVDAObject):

	def getChildID(self):
		return int(re.match(r'^_([0-9a-f]*)_.*$',self.accObject.ia.this).group(1),16)*-1

	def event_focusObject(self):
		self.speakObject()
		api.setVirtualBufferCursor(api.getVirtualBuffer().getCaretIndex())

class NVDAObject_mozillaContentWindowClass_document(NVDAObject_mozillaContentWindowClass):

	def getValue(self):
		return ""

	def event_focusObject(self):
		if self==api.getVirtualBuffer().getRootObject():
			return
		audio.speakMessage("Loading document...")
		if not self.getStates()&STATE_SYSTEM_BUSY:
			api.setVirtualBuffer(self.getWindowHandle())
			api.setVirtualBufferCursor(api.getVirtualBuffer().getCaretIndex())
			audio.cancel()
			audio.speakText(api.getVirtualBuffer().getText())

	def event_objectStateChange(self):
		states=self.getStates()
		states_on=states-(states&self.lastStates)
		states_off=self.lastStates-(states&self.lastStates)
		debug.writeMessage("doc state: %s, %s"%(getStateNames(states_on),getStateNames(states_off)))
		if states_on&STATE_SYSTEM_BUSY:
			audio.speakMessage("Loading document...")
		if states_off&STATE_SYSTEM_BUSY:
			api.setVirtualBuffer(self.getWindowHandle())
			audio.cancel()
			audio.speakText(api.getVirtualBuffer().getText())
		NVDAObject_mozillaContentWindowClass(self)

class NVDAObject_mozillaContentWindowClass_link(NVDAObject_mozillaContentWindowClass):

	def getValue(self):
		return ""

	def filterStates(self,states):
		states=NVDAObject_mozillaContentWindowClass.filterStates(self,states)
		states-=(states&STATE_SYSTEM_LINKED)
		states-=(states&STATE_SYSTEM_TRAVERSED)
		return states

	def getTypeString(self):
		states=self.getStates()
		typeString=""
		if states&STATE_SYSTEM_TRAVERSED:
			typeString+="visited "
		if states&STATE_SYSTEM_SELECTED:
			typeString+="same page "
		typeString+=NVDAObject.getTypeString(self)
		return typeString

	def getChildren(self):
		children=NVDAObject.getChildren(self)
		if (len(children)==1) and (NVDAObject.getRole(children[0])==pyAA.Constants.ROLE_SYSTEM_TEXT):
			return []
		return children

class NVDAObject_mozillaContentWindowClass_listItem(NVDAObject_mozillaContentWindowClass):

	def getName(self):
		child=self.getFirstChild()
		if child and self.getRole()==ROLE_SYSTEM_STATICTEXT:
			name=child.getName()
		else:
			name=""
		return name

	def getChildren(self):
		children=NVDAObject.getChildren(self)
		if (len(children)>=1) and (NVDAObject.getRole(children[0])==ROLE_SYSTEM_STATICTEXT):
			del children[0]
		return children

class NVDAObject_mozillaContentWindowClass_text(NVDAObject_mozillaContentWindowClass):

	def getName(self):
		name=NVDAObject.getName(self)
		value=NVDAObject.getValue(self)
		if name and not value:
			return ""
		else:
			return name

	def getRole(self):
		if NVDAObject.getStates(self)&STATE_SYSTEM_READONLY:
			return ROLE_SYSTEM_STATICTEXT
		else:
			return NVDAObject.getRole(self)

	def getValue(self):
		name=NVDAObject.getName(self)
		value=NVDAObject.getValue(self)
		if name and not value:
			return name
		else:
			return ""

class NVDAObject_TrayClockWClass(NVDAObject):

	def getRole(self):
		return ROLE_SYSTEM_CLOCK

classMap={
"Shell_TrayWnd":NVDAObject_Shell_TrayWnd,
"Progman":NVDAObject_Progman,
"#32770_18":NVDAObject_dialog,
"TrayClockWClass":NVDAObject_TrayClockWClass,
"Edit":NVDAObject_Edit,
"RICHEDIT50W":NVDAObject_Edit,
"Button_44":NVDAObject_checkBox,
"MozillaUIWindowClass":NVDAObject_mozillaUIWindowClass,
"MozillaUIWindowClass_14":NVDAObject_mozillaUIWindowClass_application,
"MozillaContentWindowClass":NVDAObject_mozillaContentWindowClass,
"MozillaContentWindowClass_15":NVDAObject_mozillaContentWindowClass_document,
"MozillaContentWindowClass_30":NVDAObject_mozillaContentWindowClass_link,
"MozillaContentWindowClass_34":NVDAObject_mozillaContentWindowClass_listItem,
"MozillaContentWindowClass_42":NVDAObject_mozillaContentWindowClass_text,
}
