# -*- coding: UTF-8 -*-
#brailleDisplayDrivers/esys.py
#A part of NonVisual Desktop Access (NVDA)
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.
#Copyright (C) 2017 NV Access Limited, Babbage B.V.

from collections import OrderedDict, defaultdict
from cStringIO import StringIO
import serial
import hwPortUtils
import braille
import inputCore
from logHandler import log
import brailleInput
import hwIo
from baseObject import AutoPropertyObject
import wx
import threading

BAUD_RATE = 9600
PARITY = serial.PARITY_EVEN

STX=b'\x02'
ETX=b'\x03'
ACK=b'\x06'
EB_SYSTEM=b'S' # 0x53
EB_MODE=b'R' # 0x52
EB_KEY=b'K' # 0x4b
EB_BRAILLE_DISPLAY=b'B' # 0x42
EB_END_KEY=b'E' # 0x45
EB_KEY_INTERACTIVE=b'I' # 0x49
EB_KEY_INTERACTIVE_SINGLE_CLICK=b'\x01'
EB_KEY_INTERACTIVE_REPETITION=b'\x02'
EB_KEY_INTERACTIVE_DOUBLE_CLICK=b'\x03'
EB_KEY_BRAILLE='B' # 0x42
EB_KEY_COMMAND=b'C' # 0x43
#EB_KEY_USB_HID_MODE=b'U' # 0x55
#EB_KEY_USB=b'u' # 0x75
EB_BRAILLE_DISPLAY_STATIC=b'S' # 0x53
EB_SYSTEM_IDENTITY=b'I' # 0x49
EB_SYSTEM_NAME=b'N' # 0x4e
EB_SYSTEM_SERIAL = 'S' # 0x53
EB_SYSTEM_BATTERY=b'B' # 0x42
EB_SYSTEM_DISPLAY_LENGTH=b'G' # 0x47
EB_SYSTEM_TYPE=b'T' # 0x54
EB_SYSTEM_OPTION=b'O'
EB_SYSTEM_SOFTWARE=b'W'
EB_SYSTEM_PROTOCOL=b'P'
EB_SYSTEM_FRAME_LENGTH=b'M'
EB_SYSTEM_DATE_AND_TIME=b'D'
EB_MODE_PILOT=b'P'
EB_MODE_INTERNAL=b'I'
EB_MODE_MENU=b'M'
EB_IRIS_TEST=b'T'
EB_IRIS_TEST_sub=b'L'
EB_VISU=b'V'
EB_VISU_DOT=b'D'

KEYS_STICK=OrderedDict({
	0x10000: "joystick1Up",
	0x20000: "joystick1Down",
	0x40000: "joystick1Right",
	0x80000: "joystick1Left",
	0x100000: "joystick1Center",    
	0x1000000: "joystick2Up",
	0x2000000: "joystick2Down",
	0x4000000: "joystick2Right",
	0x8000000: "joystick2Left",
	0x10000000: "joystick2Center"
})
KEYS_ESYS=OrderedDict({
	0x01: "switch1Right",
	0x02: "switch1Left",
	0x04: "switch4Right",
	0x08: "switch4Left",
	0x10: "switch2Right",
	0x20: "switch2Left",
	0x40: "switch3Right",
	0x80: "switch3Left",
	0x100: "switch5Right",
	0x200: "switch5Left",
	0x400: "switch6Right",
	0x800: "switch6Left",
})
KEYS_ESYS.update(KEYS_STICK)
KEYS_IRIS=OrderedDict({
	0x01: "l1",
	0x02: "l2",
	0x04: "l3",
	0x08: "l4",
	0x10: "l5",
	0x20: "l6",
	0x40: "l7",
	0x80: "l8",
	0x100: "upArrow",
	0x200: "downArrow",
	0x400: "rightArrow",
	0x800: "leftArrow",
})

KEYS_ESITIME=OrderedDict({
	0x01: "l1",
	0x02: "l2",
	0x04: "l3",
	0x08: "l4",
	0x10: "l8",
	0x20: "l7",
	0x40: "l6",
	0x80: "l5",
})
KEYS_ESITIME.update(KEYS_STICK)

DEVICE_TYPES={
	0x01:"Iris 20",
	0x02:"Iris 40",
	0x03:"Iris S20",
	0x04:"Iris S32",
	0x05:"Iris KB20",
	0x06:"IrisKB 40",
	0x07:"Esys 12",
	0x08:"Esys 40",
	0x09:"Esys Light 40",
	0x0a:"Esys 24",
	0x0b:"Esys 64",
	0x0c:"Esys 80",
	#0x0d:"Esys", # reserved in protocol
	0x0e:"Esytime 32",
	0x0f:"Esytime 32 standard",
	0x10:"Esytime evo 32",
	0x11:"Esytime evo 32 standard",
}

USB_IDS_HID = {
	"VID_C251&PID_1122", # Esys (version < 3.0, no SD card
	"VID_C251&PID_1123", # Reserved
	"VID_C251&PID_1124", # Esys (version < 3.0, with SD card
	"VID_C251&PID_1125", # Reserved
	"VID_C251&PID_1126", # Esys (version >= 3.0, no SD card
	"VID_C251&PID_1127", # Reserved
	"VID_C251&PID_1128", # Esys (version >= 3.0, with SD card
	"VID_C251&PID_1129", # Reserved
	"VID_C251&PID_112A", # Reserved
	"VID_C251&PID_112B", # Reserved
	"VID_C251&PID_112C", # Reserved
	"VID_C251&PID_112D", # Reserved
	"VID_C251&PID_112E", # Reserved
	"VID_C251&PID_112F", # Reserved
	"VID_C251&PID_1130", # Esytime
	"VID_C251&PID_1131", # Reserved
	"VID_C251&PID_1132", # Reserved
}

BLUETOOTH_NAMES = {
	"Esys",
}

def bytesToInt(bytes):
	"""Converts a basestring to its integral equivalent."""
	return int(bytes.encode('hex'), 16)

class BrailleDisplayDriver(braille.BrailleDisplayDriver):
	name = "eurobraille"
	# Translators: Names of braille displays.
	description = _("Eurobraille Esys/Esytime/Iris displays")
	isThreadSafe = True
	timeout = 0.2

	@classmethod
	def check(cls):
		return True

	@classmethod
	def getPossiblePorts(cls):
		ports = OrderedDict()
		comPorts = list(hwPortUtils.listComPorts(onlyAvailable=True))
		try:
			next(cls._getAutoPorts(comPorts))
			ports.update((cls.AUTOMATIC_PORT,))
		except StopIteration:
			pass
		for portInfo in comPorts:
			# Translators: Name of a serial communications port.
			ports[portInfo["port"]] = _("Serial: {portName}").format(portName=portInfo["friendlyName"])
		return ports

	@classmethod
	def _getAutoPorts(cls, comPorts):
		for portInfo in hwPortUtils.listHidDevices():
			if portInfo.get("usbID") in USB_IDS_HID:
				yield portInfo["devicePath"], "USB HID"
		# Try bluetooth ports last.
		for portInfo in sorted(comPorts, key=lambda item: "bluetoothName" in item):
			port = portInfo["port"]
			if "bluetoothName" in portInfo:
				# Bluetooth.
				portType = "bluetooth"
				btName = portInfo["bluetoothName"]
				if not any(btName.startswith(prefix) for prefix in BLUETOOTH_NAMES):
					continue
			else:
				continue
			yield port, portType

	def __init__(self, port="Auto"):
		super(BrailleDisplayDriver, self).__init__()
		self.numCells = 0
		self.deviceType = None
		self._deviceData = {}
		self._awaitingFrameReceipts  = {}
		self._frameLength = None
		self._frame = 0x20
		self._frameLock = threading.Lock()
		if port == "auto":
			tryPorts = self._getAutoPorts(hwPortUtils.listComPorts(onlyAvailable=True))
		else:
			tryPorts = ((port, "serial"),)
		for port, portType in tryPorts:
			# At this point, a port bound to this display has been found.
			# Try talking to the display.
			self.isHid = portType == "USB HID"
			try:
				if self.isHid:
					self._dev = hwIo.Hid(
						port,
						onReceive=self._onReceive,
						# Eurobraille wants us not to block other application's access to this handle.
						exclusiveAccess=True
					)
				else:
					self._dev = hwIo.Serial(port, baudrate=BAUD_RATE, timeout=self.timeout, writeTimeout=self.timeout, onReceive=self._onReceive)
			except EnvironmentError:
				log.debugWarning("Error while connecting to port %r"%port, exc_info=True)
				continue

			# Request device identification
			self._sendPacket(EB_SYSTEM, EB_SYSTEM_IDENTITY)
			# A device identification results in multiple packets.
			# Make sure we've received everything before we continue
			while self._dev.waitForRead(self.timeout):
				continue
			if self.numCells and self.deviceType:
				# A display responded.
				# Make sure visualisation packets are disabled, as we ignore them anyway.
				self._sendPacket(EB_VISU, EB_VISU_DOT, '0')
				log.info("Found {device} connected via {type} ({port})".format(
					device=self.deviceType, type=portType, port=port))
				break
			self._dev.close()

		else:
			raise RuntimeError("No supported Eurobraille display found")

		self.keysDown = defaultdict(int)
		self._ignoreCommandKeyReleases = False

	def terminate(self):
		try:
			super(BrailleDisplayDriver, self).terminate()
		finally:
			# Make sure the device gets closed.
			# If it doesn't, we may not be able to re-open it later.
			self._dev.close()
			self._deviceData.clear()

	def _onReceive(self, data):
		if self.isHid:
			# data contains the entire packet.
			byte0 = data[0]
			stream = StringIO(data)
			stream.seek(1)
		else:
			byte0= data
			stream = self._dev
		assert(byte0=="\x00")
		byte1=stream.read(1)
		if byte1 == ACK:
			frame=ord(stream.read(1))
			self._handleAck(frame)
		elif byte1 == STX:
			length = bytesToInt(stream.read(2))-2 # lenght includes the lenght itself
			packet = stream.read(length)
			assert(stream.read(1)==ETX)
			packetType = packet[0]
			packetSubType = packet[1]
			packetData = packet[2:] if length>2 else ""
			if packetType==EB_SYSTEM:
				self._handleSystemPacket(packetSubType, packetData)
			elif packetType==EB_MODE and packetSubType  == EB_MODE_PILOT:
				# This packet means the display is returning from internal mode
				# Rewrite the current display content
				braille.handler.update()
			elif packetType==EB_KEY:
				self._handleKeyPacket(packetSubType, packetData)
			elif packetType==EB_IRIS_TEST and packetSubType==EB_IRIS_TEST_sub:
				# Ping command sent by Iris every two seconds, send it back on the main thread.
				wx.CallAfter(self._sendPacket, packetType, packetSubType, packetData)
			elif packetType==EB_VISU:
				log.debug("Ignoring visualisation packet")
			else:
				log.debug("Ignoring packet: type %s, subtype %s, data %s"%(
					packetType,
					packetSubType,
					packetData
				))

	def _handleAck(self, frame):
		try:
			del self._awaitingFrameReceipts[frame]
		except KeyError:
			log.debugWarning("Received ACK for unregistered frame %d"%frame)
		super(BrailleDisplayDriver,self)._handleAck()

	def _handleSystemPacket(self, type, data):
		if type==EB_SYSTEM_TYPE:
			deviceType=ord(data)
			self.deviceType = DEVICE_TYPES[deviceType]
			if 0x01<=deviceType<=0x06: # Iris
				self.keys=KEYS_IRIS
			elif 0x07<=deviceType<=0x0d: # Esys
				self.keys=KEYS_ESYS
			elif 0x0e<=deviceType<=0x11: # Esitime
				self.keys=KEYS_ESITIME
			else:
				log.debugWarning("Unknown device identifier %s"%data)
		elif type==EB_SYSTEM_DISPLAY_LENGTH:
			self.numCells=ord(data)
		elif type==EB_SYSTEM_FRAME_LENGTH:
			self._frameLength=bytesToInt(data)
		elif type==EB_SYSTEM_PROTOCOL:
			protocol=data.rstrip("\x00 ")
			try:
				version=float(protocol[:4])
			except ValueError:
				pass
			else:
				self.receivesAckPackets = version>=3.0
		elif type==EB_SYSTEM_IDENTITY:
			return # End of system information
		self._deviceData[type]=data.rstrip("\x00 ")

	def _handleKeyPacket(self, group, data):
		arg = bytesToInt(data)
		if group == EB_KEY_INTERACTIVE and data[0]==EB_KEY_INTERACTIVE_REPETITION:
			log.debug("Ignoring routing key %d repetition"%(ord(data[1])-1))
			return
		if arg==self.keysDown[group]:
			log.debug("Ignoring key repetition")
			return
		self.keysDown[group] |= arg
		if group == EB_KEY_COMMAND and arg>=self.keysDown[group]:
			# Started a gesture including command keys
			self._ignoreCommandKeyReleases = False
		else:
			if group != EB_KEY_COMMAND or not self._ignoreCommandKeyReleases:
				try:
					inputCore.manager.executeGesture(InputGesture(self))
				except inputCore.NoInputGestureAction:
					pass
				self._ignoreCommandKeyReleases = group == EB_KEY_COMMAND or self.keysDown[EB_KEY_COMMAND]>0
			if group == EB_KEY_COMMAND:
				self.keysDown[group] = arg
			else:
				del self.keysDown[group]

	def _sendPacket(self, packetType, packetSubType, packetData=""):
		packetSize=len(packetData)+4
		packet=[
			"\x00",
			STX,
			chr((packetSize>>8)&0xff),
			chr(packetSize&0xff),
			packetType,
			packetSubType,
			packetData,
			ETX
		]
		if self.receivesAckPackets:
			with self._frameLock:
				frame=self._frame
				packet.insert(-1,chr(frame))
				self._awaitingFrameReceipts[frame]=packet
				self._frame=frame+1 if frame<0x7F else 0x20
		self._dev.write("".join(packet))

	def display(self, cells):
		# cells will already be padded up to numCells.
		self._sendPacket(EB_BRAILLE_DISPLAY, EB_BRAILLE_DISPLAY_STATIC, "".join(chr(cell) for cell in cells))	

	gestureMap = inputCore.GlobalGestureMap({
		"globalCommands.GlobalCommands": {
			"braille_routeTo": ("br(eurobraille):routing","br(eurobraille):doubleRouting",),
			"braille_scrollBack": (
				"br(eurobraille):switch1Left",
				#"br(eurobraille):switch2Left",
				#"br(eurobraille):switch3Left", "br(eurobraille):switch4Left",
				#"br(eurobraille):switch5Left", "br(eurobraille):switch6Left",
			),
			"braille_scrollForward": (
				"br(eurobraille):switch1Right",
				#"br(eurobraille):switch2Right",
				#"br(eurobraille):switch3Right", "br(eurobraille):switch4Right",
				#"br(eurobraille):switch5Right", "br(eurobraille):switch6Right",
			),
			"braille_toFocus": (
				"br(eurobraille):switch1Left+switch1Right", "br(eurobraille):switch2Left+switch2Right",
				"br(eurobraille):switch3Left+switch3Right", "br(eurobraille):switch4Left+switch4Right",
				"br(eurobraille):switch5Left+switch5Right", "br(eurobraille):switch6Left+switch6Right",
			),
			"kb:enter": ("br(eurobraille):joystick2Center",),
			"kb:upArrow": ("br(eurobraille):joystick2Up",),
			"kb:downArrow": ("br(eurobraille):joystick2Down",),
			"kb:leftArrow": ("br(eurobraille):joystick2Left",),
			"kb:rightArrow": ("br(eurobraille):joystick2Right",),
			"kb:home": ("br(eurobraille):joystick2Left+joystick2Up",),
			"kb:end": ("br(eurobraille):joystick2Right+joystick2Down",),
			"review_previousLine": ("br(eurobraille):joystick1Up",),
			"review_nextLine": ("br(eurobraille):joystick1Down",),
			"review_previousCharacter": ("br(eurobraille):joystick1Left",),
			"review_nextCharacter": ("br(eurobraille):joystick1Right",),
			"reviewMode_previous": ("br(eurobraille):joystick1Left+joystick1Up",),
			"reviewMode_next": ("br(eurobraille):joystick1Right+joystick1Down",),
			# Esys has a dedicated key for backspace and combines backspace and space to perform a return.
			"braille_eraseLastCell": ("br(eurobraille):backSpace",),
			"braille_enter": ("br(eurobraille):backSpace+space",),
		},
	})

class InputGesture(braille.BrailleDisplayGesture, brailleInput.BrailleInputGesture):

	source = BrailleDisplayDriver.name

	def __init__(self, display):
		super(InputGesture, self).__init__()
		self.model = display.deviceType.split(" ")[0]
		keysDown = dict(display.keysDown)
		self.keyNames = names = []
		for group, groupKeysDown in keysDown.iteritems():
			if group == EB_KEY_BRAILLE:
				if sum(keysDown.itervalues())==groupKeysDown and not groupKeysDown & 0x100:
					# This is braille input.
					# 0x1000 is backspace, 0x2000 is space
					self.dots = groupKeysDown & 0xff
					self.space = groupKeysDown & 0x200
				names.extend("dot%d" % (i+1) for i in xrange(8) if (groupKeysDown &0xff) & (1 << i))
				if groupKeysDown & 0x200:
					names.append("space")
				if groupKeysDown & 0x100:
					names.append("backSpace")
			if group == EB_KEY_INTERACTIVE: # Routing
				self.routingIndex = (groupKeysDown & 0x3f)-1
				names.append("doubleRouting" if groupKeysDown>>8 ==ord(EB_KEY_INTERACTIVE_DOUBLE_CLICK) else "routing")
			if group == EB_KEY_COMMAND:
				for key, keyName in display.keys.iteritems():
					if groupKeysDown & key:
						# This key is pressed
						names.append(keyName)

		self.id = "+".join(names)
