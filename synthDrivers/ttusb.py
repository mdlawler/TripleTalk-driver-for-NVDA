#synthDrivers/ttusbd.py
#A part of NonVisual Desktop Access (NVDA)
#copyright 2012 by NVDA contributers etc
#www.nvda-project.org/
#TripleTalk USB Driver
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

import threading
import winAPI
from winAPI import secureDesktop
import api
import synthDriverHandler
from synthDriverHandler import SynthDriver, synthDoneSpeaking, synthIndexReached 
from speech.commands import IndexCommand, PitchCommand, CharacterModeCommand
from collections import OrderedDict
from logHandler import log
import ctypes
import _ctypes
from ctypes import *
import os
import time
from autoSettingsUtils.driverSetting import DriverSetting
from autoSettingsUtils.utils import StringParameterInfo
kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
kernel32.GetPrivateProfileIntW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.INT, wintypes.LPCWSTR]
kernel32.WritePrivateProfileStringW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR]
dll_file_name = r"\ttusbd.dll"
USBTT = None
changedDesktop = False
settingPauseMode = False
lastSentIndex = 0
lastReceivedIndex = 0
stopIndexing = False
indexReached = None
indexesAvailable = None
milliseconds = 100
nvdaIndexes = [0] * 100
variants = {
	0:"Perfect Paul",
	1:"Vader",
	2:"Mountain Mike",
	3:"Precise Pete",
	4:"Jammin Jimmy",
	5:"Biff",
	6:"Skip",
	7:"Robo Robert" }
pauseModes={
	"0": StringParameterInfo("0", _("Normal")),
	"1": StringParameterInfo("1", _("No Pauses")),
	"2": StringParameterInfo("2", _("Shorten more than normal")),
	"3": StringParameterInfo("3", _("Shorten maximum amount")) }

def is_admin():
	try:
		return ctypes.windll.shell32.IsUserAnAdmin()
	except:
		return False

def unload_dll():
	global USBTT
	global lastSentIndex
	global lastReceivedIndex
	global nvdaIndexes
	if USBTT:
		_ctypes.FreeLibrary(USBTT._handle)
		USBTT = None
		nvdaIndexes.clear()
		lastSentIndex = 0
		lastReceivedIndex = 0

def load_dll(dll_name):
	global nvdaIndexes
	global USBTT
	if not USBTT:
		path = os.getenv('windir', r"c:\windows")
		dll_path = path + dll_name
		USBTT = cdll.LoadLibrary(dll_path)
		nvdaIndexes = [0] * 100

def desktopChanged(isSecureDesktop):
	# the TT dll is unhappy if it is loaded by more than one program so unload and reload when the desktop changes since secure desktops use a second copy of NVDA.
	global changedDesktop
	global settingPauseMode
	changedDesktop = True
	if isSecureDesktop:
		unload_dll()
	else:
		if settingPauseMode:
			settingPauseMode = False
			time.sleep(1) # since shellexecute returns immediately we need to give powershell time to execute the script
		load_dll(dll_file_name)

class IndexingThread(threading.Thread):
	def run(self):
		global lastReceivedIndex
		# The TT uses indexes 0-99 so we map the NVDA indexes to this and when we receive a TT index we send back the correct NVDA index
		while not stopIndexing:
			if lastSentIndex-1 == lastReceivedIndex or not USBTT:
				indexesAvailable.clear()
				indexesAvailable.wait()
			time.sleep(milliseconds/1000)
			if USBTT:
				b = USBTT.USBTT_ReadByte()
			else:
				b = -1
			if b != -1:
				lastReceivedIndex = b
				indexReached(nvdaIndexes[lastReceivedIndex])

class SynthDriver(synthDriverHandler.SynthDriver):
	name="ttusb"
	description=_("TripleTalk USB")
	supportedSettings=(
		SynthDriver.RateSetting(10),
		SynthDriver.PitchSetting(10),
		SynthDriver.InflectionSetting(10),
		SynthDriver.VolumeSetting(10),
		SynthDriver.VariantSetting(),
		DriverSetting("pauseMode", _("&Pauses"), defaultVal="0")
	)
	supportedCommands = {
		IndexCommand,
		PitchCommand,
		CharacterModeCommand
	}
	supportedNotifications = {synthIndexReached, synthDoneSpeaking}
	minRate = 0
	maxRate = 9
	minPitch=0
	maxPitch=99
	minInflection=0
	maxInflection=9
	minVolume = 0
	maxVolume = 9

	@classmethod
	def check(cls):
		return True

	def __init__(self):
		self.tt_pauseMode = kernel32.GetPrivateProfileIntW("ttalk_usb_comm", "nopauses", 0, "ttusbd.ini")
		load_dll(dll_file_name)
		if USBTT:
			init_string = b"\x18\x1e\x017b\r"
			for element in init_string:
				if element == 0x1e:
					USBTT.USBTT_WriteByteImmediate(element)
				else:
					USBTT.USBTT_WriteByte(element)
		super(synthDriverHandler.SynthDriver,self).__init__()
		global indexReached
		global indexesAvailable
		global lastSentIndex
		global lastReceivedIndex
		indexReached = self.onIndexReached
		lastSentIndex = 0
		lastReceivedIndex = 0
		indexesAvailable = threading.Event()
		indexesAvailable.clear()
		self.indexingThread = IndexingThread()
		self.indexingThread.start()
		self.tt_rate = 4
		self.tt_rateChanged = False
		self.nvda_rate = 40
		self.tt_pitch=50
		self.tt_pitchChanged = False
		self.capPitch = False
		self.tt_inflection=5
		self.tt_inflectionChanged = False
		self.nvda_inflection = 50
		self.tt_volume = 5
		self.tt_volumeChanged = False
		self.nvda_volume = 50
		self.tt_variant = "0"
		self.tt_variantChanged = False
		if not api.getForegroundObject()  == None:
			self.lastForegroundWindowHandle = api.getForegroundObject().windowHandle
		else:
			self.lastForegroundWindowHandle =0 
		winAPI.secureDesktop.post_secureDesktopStateChange.register(desktopChanged)

	def speak(self, speechSequence):
		self.pause(False) # the TripleTalk needs to be told to resume it doesn't do it upon receiving new speech and NVDA doesn't send a pause False command before sending new speech
		if not USBTT:
			return
		global lastSentIndex
		global nvdaIndexes
		global milliseconds
		global changedDesktop
		text_list = []
		item_list = []
		characterMode = False
		textLength = 0
		allowOClock = True
		hasSeconds = False
		leadingZero = False
		itemIndex = 0
		moneyString = ""
		params = b""
		totalItems = len(speechSequence) - 1
		itemLen = 0
		upperAscii = {
			128:"euro",
			129:"",
			130:"single low-9 quote",
			131:"f hook",
			132:"double low-9 quote",
			133:"horizontal ellipsis",
			134:"dagger",
			135:"double dagger",
			136:"circumflex accent",
			137:"per mille",
			138:"S caron",
			139:"single left-pointing angle quote",
			140:"ligature OE",
			141:"",
			142:"Z caron",
			143:"",
			144:"",
			145:"left single quote",
			146:"right single quote",
			147:"left double quote",
			148:"right double quote",
			149:"bullet",
			150:"en dash",
			151:"em dash",
			152:"tilde",
			153:"trade mark",
			154:"S caron",
			155:"single right-pointing angle quote",
			156:"ligature oe",
			157:"",
			158:"z caron",
			159:"Y diaeresis",
			160:"non-breaking space",
			161:"inverted exclamation mark",
			162:"cents",
			163:"pound",
			164:"currency",
			165:"yen",
			166:"pipe broken vertical bar",
			167:"section",
			168:"spacing diaeresis – umlaut",
			169:"copyright",
			170:"feminine ordinal",
			171:"left double-angle quotes",
			172:"egation",
			173:"soft hyphen",
			174:"registered trademark",
			175:"spacing macron – overline",
			176:"degree",
			177:"plus-or-minus",
			178:"superscript two",
			179:"superscript three",
			180:"acute accent",
			181:"micro",
			182:"pilcrow",
			183:"middle dot",
			184:"spacing cedilla",
			185:"superscript one",
			186:"masculine ordinal",
			187:"right double-angle quotes",
			188:"one quarter",
			189:"one half",
			190:"three quarters",
			191:"inverted question mark",
			192:"A grave",
			193:"A acute",
			194:"A circumflex",
			195:"A tilde",
			196:"A diaeresis",
			197:"A ring above",
			198:"A E",
			199:"C cedilla",
			200:"E grave",
			201:"E acute",
			202:"E circumflex",
			203:"E diaeresis",
			204:"I grave",
			205:"I acute",
			206:"I circumflex",
			207:"I diaeresis",
			208:"ETH",
			209:"N tilde",
			210:"O grave",
			211:"O acute",
			212:"O circumflex",
			213:"O tilde",
			214:"O diaeresis",
			215:"times",
			216:"O a slash",
			217:"U grave",
			218:"U acute",
			219:"U circumflex",
			220:"U diaeresis",
			221:"Y acute",
			222:"THORN",
			223:"sharp s",
			224:"a grave",
			225:"a acute",
			226:"a circumflex",
			227:"a tilde",
			228:"a diaeresis",
			229:"a ring above",
			230:"a e",
			231:"c cedilla",
			232:"e grave",
			233:"e acute",
			234:"e circumflex",
			235:"e diaeresis",
			236:"i grave",
			237:"i acute",
			238:"i circumflex",
			239:"i diaeresis",
			240:"eth",
			241:"n tilde",
			242:"o grave",
			243:"o acute",
			244:"o circumflex",
			245:"o tilde",
			246:"o dia,eresis",
			247:"divided",
			248:"o slash",
			249:"u grave",
			250:"u acute",
			251:"u circumflex",
			252:"u diaeresis",
			253:"y acute",
			254:"thorn,",
			255:"y diaeresis" }
		for item in speechSequence:
			if isinstance(item, CharacterModeCommand):
				characterMode = item.state
			elif isinstance(item, str):
				item_list.clear()
				allowOClock = True
				leadingZero = False
				moneyString = ""
				itemIndex = 0
				itemLen = len(item)
				for elementIndex, element in enumerate(item):
					# don't allow the text to contain the control character control-A as this could put the synth in a strange state
					if ord(element) == 1:
						if not item_list: item_list = list(item)
						item_list[elementIndex] = " "
					if not characterMode:
						#fix so that the synthesizer says time correctly and also make it report leading zeros for numbers.
						# remove the comma character only when it is in a number so the synthesizer says numbers correctly
						# this only seems to matter if you set nopauses=1 in ttusbd.ini in the windows directory
						# fix issues the synth has pronouncing money
						if elementIndex < itemIndex: # skip the indexes we already processed for point or money the previous time through the loop
							continue
						itemIndex = 0
						if element.isnumeric() or element == ':' or element == '$' or element == ',' or element == '.':
							if element == '.': # make it pronounce decimals correctly
								if elementIndex == 0 or (elementIndex > 0 and (item[elementIndex-1].isnumeric() or item[elementIndex-1].isspace())) and elementIndex+1 in range(itemLen) and item[elementIndex+1].isnumeric:
									if not item_list: item_list = list(item)
									item_list[elementIndex] = " point "
								itemIndex = elementIndex+1
								while itemIndex in range(itemLen) and item[itemIndex] == '0':
									if not item_list: item_list = list(item)
									item_list[itemIndex] = "o "
									itemIndex+=1
							elif element == ',':
								if elementIndex > 0 and item[elementIndex-1].isnumeric() and elementIndex+1 in range(itemLen) and item[elementIndex+1].isnumeric:
									if not item_list: item_list = list(item)
									item_list[elementIndex] = ""
							elif element == '$':
								if elementIndex > 0 and not item[elementIndex-1].isspace():
									if not item_list: item_list = list(item)
									item_list[elementIndex] = " $"
								itemIndex = elementIndex+1
								while itemIndex in range(itemLen):
									if not item[itemIndex].isnumeric() and item[itemIndex] != '.':
										break
									elif item[itemIndex] == '.':
										if itemIndex+3 in range(itemLen) and item[itemIndex+1].isnumeric() and item[itemIndex+2].isnumeric() and item[itemIndex+3].isnumeric():
											tempIndex = itemIndex+1
											while tempIndex in range(itemLen) and item[tempIndex].isnumeric():
												tempIndex+=1
											moneyString = item[tempIndex-1]
											moneyString += " cents "
											if not item_list: item_list = list(item)
											item_list[tempIndex-1] = moneyString
										elif itemIndex+2 in range(itemLen) and item[itemIndex+1] == '0' and (item[itemIndex+2] == '0' or not item[itemIndex+2].isnumeric()):
											if not item_list: item_list = list(item)
											item_list[itemIndex+1] = ""
											item_list[itemIndex+2] = ""
										elif itemIndex+2 in range(itemLen) and item[itemIndex+1].isnumeric() and item[itemIndex+1] > '0' and not item[itemIndex+2].isnumeric():
											if not item_list: item_list = list(item)
											moneyString = " and "
											moneyString += item_list[itemIndex+1]
											moneyString += "0 cents "
											item_list[itemIndex+1] = moneyString
										elif itemIndex+2 in range(itemLen) and item[itemIndex+1] == '0' and item[itemIndex+2] == '1':
											if not item_list: item_list = list(item)
											item_list[itemIndex+1] = "and 1 cent"
											item_list[itemIndex+2] = ""
										elif itemIndex+2 in range(itemLen) and item[itemIndex+1].isnumeric() and item[itemIndex+2].isnumeric():
											if item[itemIndex+1] == '0':
												moneyString = " and "
												moneyString += item[itemIndex+2]
												moneyString += " cents "
											elif item[itemIndex+1] >= '1':
												moneyString = " and "
												moneyString += item[itemIndex+1]
												moneyString += item[itemIndex+2]
												moneyString += " cents "
											if not item_list: item_list = list(item)
											item_list[itemIndex+1] = moneyString
											item_list[itemIndex+2] = ""
									itemIndex+=1
							elif element == ':':
								if elementIndex >= 3 and item[elementIndex-3] == ':':
									hasSeconds = True
								if elementIndex >= 2 and item[elementIndex-1] == '0' and item[elementIndex-2] == '0':
									if not hasSeconds and item_list and item_list[elementIndex-2] != 'o ':
										if not item_list: item_list = list(item)
										item_list[elementIndex	-2] = "zero "
								if elementIndex == 0 or item[elementIndex-1] == '0':
									allowOClock = False
								if elementIndex >= 2 and item[elementIndex-2] == '1': #for 10 o clock
									allowOClock = True
							if element == ':' and elementIndex+2 in range(itemLen) and item[elementIndex+1] == '0':
								if not item_list: item_list = list(item)
								if allowOClock and not hasSeconds:
									if item[elementIndex+2].isnumeric():
										item_list[elementIndex+1] = "o "
								else:
									item_list[elementIndex+1] = "zero "
								if item[elementIndex+2] == '0':
									if not item_list: item_list = list(item)
									if allowOClock and not hasSeconds:
										item_list[elementIndex+2] = "clock "
									else:
										item_list[elementIndex+2] = "zero "
							elif element == '0':
								if elementIndex == 0:
									leadingZero = True
								elif elementIndex > 0 and item[elementIndex-1].isspace() and elementIndex+1 in range(itemLen) and not item[elementIndex+1].isspace():
									leadingZero = True
								itemIndex = elementIndex
								while leadingZero and itemIndex in range(itemLen):
									if item[itemIndex] == ':':
										leadingZero = False
									elif not item[itemIndex].isnumeric():
										break
									itemIndex +=1
								itemIndex = 0 # set it back to 0 since we use it to skip processed money elements when it is > 0
								if leadingZero:
									if not item_list: item_list = list(item)
									item_list[elementIndex] = "zero "
							else:
								leadingZero = False
								moneyString = ""
								allowOClock = True
								hasSeconds = False
				if item_list:
					item = "".join(item_list)
				if characterMode and len(item) == 1 and ord(item) > 127:
					item = item.replace(item, upperAscii[ord(item)])
				text_list.append(item)
				# when NVDA sends shortcut characters such as alt n it doesn't put a space after the shortcut and this synthesizer needs that to not run it together the next word
				if characterMode:
						text_list.append(" ")
			elif isinstance(item, IndexCommand):
				text_list.append("\x1e\x01%di" % lastSentIndex)
				nvdaIndexes[lastSentIndex] = item.index
				lastSentIndex += 1
				if lastSentIndex == 100:
					lastSentIndex = 0
			elif isinstance(item, PitchCommand):
				self.capPitch = True
				offsetPitch = self.tt_pitch + item.offset
				if offsetPitch > self.maxPitch:
					offsetPitch = self.maxPitch
				text_list.append("\x1e\x01%dp" % offsetPitch)

		text = "".join(text_list).replace("\r", "")
		text = text.encode('ascii', 'replace')
		textLength = len(text)
		# only resend the speech parameters when the foreground window changes or when the desktop changes.
		# This accounts for self talking apps that might change the speech parameters and the version of NVDA running on the secure desktop doing the same
		# force this by lying and saying that the variant of the voice has changed because when it changes all parameters have to be resent
		if self.lastForegroundWindowHandle != api.getForegroundObject().windowHandle:
			self.tt_variantChanged = True
			self.lastForegroundWindowHandle = api.getForegroundObject().windowHandle
		if changedDesktop:
			self.tt_variantChanged = True
			changedDesktop = False
		if self.tt_variantChanged:
			params = ("\x1e\x01%so\x01%ds\x01%dp\x01%de\x01%dv" % (self.tt_variant, self.tt_rate, self.tt_pitch, self.tt_inflection, self.tt_volume)).encode('ascii', 'replace')
			self.tt_variantChanged = False
			self.tt_rateChanged = False
			self.tt_pitchChanged = False
			self.tt_volumeChanged = False
			self.tt_inflectionChanged = False
		if self.tt_rateChanged:
			params += ("\x1e\x01%ds" % self.tt_rate).encode('ascii', 'replace')
			self.tt_rateChanged = False
		if self.tt_pitchChanged:
			params += ("\x1e\x01%dp" % self.tt_pitch).encode('ascii', 'replace')
			self.tt_pitchChanged = False
		if self.tt_volumeChanged:
			params += ("\x1e\x01%dv" % self.tt_volume).encode('ascii', 'replace')
			self.tt_volumeChanged = False
		if self.tt_inflectionChanged:
			params += ("\x1e\x01%de" % self.tt_inflection).encode('ascii', 'replace')
			self.tt_inflectionChanged = False
		text = b"%s %s" % (params, text)
		text = text + b"\r"
		# don't use WriteString because it has performance issues, causes other strange behavior, and is just meant for quick testing
		for element in text:
			if element == 0x1e:
				USBTT.USBTT_WriteByteImmediate(element)
			else:
				USBTT.USBTT_WriteByte(element)
		indexesAvailable.set()
		if characterMode or textLength < 10:
			milliseconds = 10 # for short strings use 10 milliseconds to keep things responsive
		else:
			milliseconds = 100 # for long strings use 100 milliseconds as to not hammer the synth for index marks and waste CPU
		# Sometimes the synth can get stuck with the cap pitch offset this forces it to reset pitch after changing pitch offset for caps
		if self.capPitch:
			self.tt_pitchChanged = True
			self.capPitch = False

	def cancel(self):
		self.pause(False) # the TripleTalk needs to be told to resume it doesn't do it upon receiving new speech and NVDA doesn't send a pause False command before sending new speech
		if not USBTT:
			return
		USBTT.USBTT_WriteByteImmediate(0x18)

	def terminate(self):
		global indexReached
		global stopIndexing
		unload_dll()
		indexReached = None
		stopIndexing = True

	def _set_rate(self, rate):
		if rate != self.tt_rate:
			self.tt_rateChanged = True
			self.tt_rate = round(rate/10)
			self.nvda_rate = rate
			if self.tt_rate > self.maxRate:
				self.tt_rate = self.maxRate

	def _get_rate(self):
		return self.nvda_rate

	def _set_pitch(self, pitch):
		if pitch != self.tt_pitch:
			self.tt_pitchChanged = True
			val = round(self._percentToParam(pitch, self.minPitch, self.maxPitch), -1)
			if val > self.maxPitch:
				val = self.maxPitch
			self.tt_pitch = val

	def _get_pitch(self):
		return round(self._paramToPercent(self.tt_pitch, self.minPitch, self.maxPitch), -1)

	def _set_inflection(self, inflection):
		if inflection != self.tt_inflection:
			self.tt_inflectionChanged = True
			self.tt_inflection = round(inflection/10)
			self.nvda_inflection = inflection
			if self.tt_inflection > self.maxInflection:
				self.tt_inflection = self.maxInflection

	def _get_inflection(self):
		return self.nvda_inflection

	def _set_volume(self, volume):
		if volume != self.tt_volume:
			self.tt_volumeChanged = True
			self.tt_volume = round(volume/10)
			self.nvda_volume = volume
			if self.tt_volume > self.maxVolume:
				self.tt_volume = self.maxVolume

	def _get_volume(self):
		return self.nvda_volume

	def _getAvailableVariants(self):
		return OrderedDict((str(id), synthDriverHandler.VoiceInfo(str(id), name)) for id, name in variants.items())

	def _set_variant(self, v):
		if v != self.tt_variant:
			self.tt_variantChanged = True
			self.tt_variant = v if int(v) in variants else "0"

	def _get_variant(self):
		return self.tt_variant

	def onIndexReached(self, index):
		if index >= 0:
			synthIndexReached.notify(synth=self, index=index)
		else:
			synthDoneSpeaking.notify(synth=self)

	def pause(self,switch):
		if not USBTT:
			return
		if switch:
			USBTT.USBTT_WriteByteImmediate(0x10)
		else:
			USBTT.USBTT_WriteByteImmediate(0x12)

	def _get_availablePausemodes(self):
		return pauseModes

	def _set_pauseMode(self, val):
		global settingPauseMode
		result = 0
		writeVal = val
		val = int(val)	
		if val != self.tt_pauseMode:
			if is_admin():
				result = kernel32.WritePrivateProfileStringW("ttalk_usb_comm", "nopauses", writeVal, "ttusbd.ini")
			else:
				args = " -file "
				args += os.path.dirname(__file__)
				args += r"\ttpausecontrol.ps1 "
				args += writeVal
				result = ctypes.windll.shell32.ShellExecuteW(None, "runas", "powershell", args, None, 0)
				if result < 32:
					result = 0
			if result:
				if is_admin():
					unload_dll()
					load_dll(dll_file_name)
				else:
					settingPauseMode = True
				self.tt_pauseMode = val

	def _get_pauseMode(self):
		return str(self.tt_pauseMode)
