#synthDrivers/ttusbd.py
#A part of NonVisual Desktop Access (NVDA)
#copyright 2012 by NVDA contributers etc
#www.nvda-project.org/
#TripleTalk USB Driver
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

from inspect import currentframe, getframeinfo
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
import sys
import time
from autoSettingsUtils.driverSetting import DriverSetting
from autoSettingsUtils.utils import StringParameterInfo
kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
kernel32.GetPrivateProfileIntW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.INT, wintypes.LPCWSTR]
kernel32.WritePrivateProfileStringW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR]
USBTT = None
exceptionLine = 0
changedDesktop = False
settingPauseMode = False
lastSentIndex = 0
lastReceivedIndex = -1
stopIndexing = False
indexReached = None
indexesAvailable = None
milliseconds = 100
nvdaIndexes = [0] * 100
synthFlushed = True

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
		lastReceivedIndex = -1

def load_dll(load):
	global nvdaIndexes
	global USBTT
	global exceptionLine
	if not USBTT:
		path = os.getenv('windir', r"c:\windows")
		path += r"\ttusbd.dll"
		if os.path.exists(path):
			if not load:
				return True
			USBTT = cdll.LoadLibrary(path)
			if USBTT:
				if not callable(getattr(USBTT, 'USBTT_WriteByte', None)):
					USBTT = None
					frameinfo = getframeinfo(currentframe())
					exceptionLine = frameinfo.lineno
				if not callable(getattr(USBTT, 'USBTT_WriteByteImmediate', None)):
					USBTT = None
					frameinfo = getframeinfo(currentframe())
					exceptionLine = frameinfo.lineno
				if not callable(getattr(USBTT, 'USBTT_ReadByte', None)):
					USBTT = None
					frameinfo = getframeinfo(currentframe())
					exceptionLine = frameinfo.lineno
				nvdaIndexes = [0] * 100
				return True
			else:
				frameinfo = getframeinfo(currentframe())
				exceptionLine = frameinfo.lineno
				return False
		else:
			frameinfo = getframeinfo(currentframe())
			exceptionLine = frameinfo.lineno
			return False
	else:
		return True

def desktopChanged(isSecureDesktop):
	# don't do this if we are running in the slave process as we never want to do unload or load in that process
	if "slave" in sys.executable.casefold():
		return
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
		load_dll(True)

class IndexingThread(threading.Thread):
	def __init__(self):
		threading.Thread.__init__(self)
		self.daemon = True
	def run(self):
		global lastReceivedIndex
		global synthFlushed
		# The TT uses indexes 0-99 so we map the NVDA indexes to this and when we receive a TT index we send back the correct NVDA index
#		log.warning("index thread %d" % threading.current_thread().ident) # uncomment this to get the indexing thread it in the nvda log for performance profiling
		while not stopIndexing:
			if lastSentIndex-1 == lastReceivedIndex:
				# we've received all sent indexes so set synthFlushed to True to notify the index thread to wait and the cancel routine to no longer send flushes until more text is being sent.
				synthFlushed = True
			if synthFlushed or not USBTT:
				indexesAvailable.clear()
				indexesAvailable.wait()
			time.sleep(milliseconds/1000)
			if synthFlushed: # when the synth has been flushed there is no need to waste time looking for indexes
				continue
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

	@classmethod
	def check(cls):
		return load_dll(False)
	def __init__(self):
		self.minRate = 0
		self.maxRate = 9
		self.minPitch=0
		self.maxPitch=99
		self.minInflection=0
		self.maxInflection=9
		self.minVolume = 0
		self.maxVolume = 9
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
		self.pauseModeOn = False
		if not api.getForegroundObject() == None:
			self.lastForegroundProcessID = api.getForegroundObject().processID
		else:
			self.lastForegroundProcessID =0 
		self.previousMetric = False
		self.tt_pauseMode = kernel32.GetPrivateProfileIntW("ttalk_usb_comm", "nopauses", 0, "ttusbd.ini")
		load_dll(True)
		if USBTT:
			init_string = b"\x18\x1e\x017b\r"
			for element in init_string:
				if element == 0x1e:
					USBTT.USBTT_WriteByteImmediate(element)
				else:
					USBTT.USBTT_WriteByte(element)
		else:
			sys.tracebacklimit = 0
			raise RuntimeError("No TripleTalk drivers available problem line %d" % exceptionLine)
		global stopIndexing
		global indexesAvailable
		global lastSentIndex
		global lastReceivedIndex
		global indexReached
		stopIndexing = False
		lastSentIndex = 0
		lastReceivedIndex = -1
		indexesAvailable = threading.Event()
		indexesAvailable.clear()
		indexReached = self.onIndexReached
		self.indexingThread = IndexingThread()
		self.indexingThread.start()
		winAPI.secureDesktop.post_secureDesktopStateChange.register(desktopChanged)
		super(synthDriverHandler.SynthDriver,self).__init__()

	def speak(self, speechSequence):
		if self.pauseModeOn:
			self.pause(False) # the TripleTalk needs to be told to resume it doesn't do it upon receiving new speech and NVDA doesn't send a pause False command before sending new speech
		if not USBTT:
			return
		global milliseconds
		global changedDesktop
		text_list = []
		characterMode = False
		for item in speechSequence:
			if isinstance(item, CharacterModeCommand):
				characterMode = item.state
			elif isinstance(item, str):
				item_list = []
				itemIndex = 0
				itemLen = len(item)
				if self.previousMetric:
					# the previous string ended with a number so put a space at the beginning of this one as not to run numbers together	
					# We have to do it this way instead of just adding a space at the end of the previous string while processing it
					# because programs like excel can send something like "28  " "l 16  " in two separate strings
					# and our metric processing would turn it into "28" "l 16" blocking the metric processing, but if the next string
					# starts with a number we have to prefix it with a space
					# if we do it while processing the previous string we'll get "28 " "l 16 " which won't block the metric processing
					self.previousMetric = False
					if itemLen and not characterMode and item[0].isnumeric():
						if not item_list: item_list = list(item)
						metricString = " "
						metricString += item[0]
						item_list[0] = metricString
				for elementIndex, element in enumerate(item):
					# Block everything below 32.  The TT dll does some of this, but it can't block the control char 0x01, flush 0x18, etc and allowing these in the text would cause problems with the synth.
					if ord(element) < 32:
						if not item_list: item_list = list(item)
						item_list[elementIndex] = " "
					elif not characterMode:
						#fix so that the synthesizer says time correctly and also make it report leading zeros for numbers.
						# remove the comma character only when it is in a number so the synthesizer says numbers correctly
						# this only seems to matter if you set nopauses=1 in ttusbd.ini in the windows directory
						# fix issues the synth has pronouncing money
						# stop the synth from saying metric stuff like liters, grams, etc
						# fix date stuff like 1st, 2nd, 3rd, 4th, etc
						if elementIndex < itemIndex: # skip the indexes we already processed for point, leading zeros, money, metric stuff, or date stuff the previous time through the loop
							continue
						itemIndex = 0
						if element == 'M' and elementIndex+2 in range(itemLen) and item[elementIndex+1] == 'c' and item[elementIndex+2] == ' ':
							# NVDA splitting words of mixed case is bad for things like McDonalds so prevent it
							if not item_list: item_list = list(item)
							item_list[elementIndex+2] = ""
						elif (((element == 'n' or element == 'N' or element == 'r' or element == 'R') and elementIndex+1 in range(itemLen) and (item[elementIndex+1] == 'd' or item[elementIndex+1] == 'D')) or
							((element == 's' or element == 'S') and elementIndex+1 in range(itemLen) and (item[elementIndex+1] == 't' or item[elementIndex+1] == 'T')) or
							((element == 't' or element == 'T') and elementIndex+1 in range(itemLen) and (item[elementIndex+1] == 'h' or item[elementIndex+1] == 'H')
							and elementIndex+2 in range(itemLen) and item[elementIndex+2] == ' ')):
							# prevent the synth from doing date stuff like 21st being twenty first instead of 21 st etc.
							tempIndex = elementIndex-1
							while tempIndex >= 0:
								if not item[tempIndex] == ' ':
									break
								tempIndex-=1
							if tempIndex>= 0 and item[tempIndex].isnumeric():
								if element == 's' or element == 'S':
									if not item_list: item_list = list(item)
									tempString = item[tempIndex]
									tempString += ","
									item_list[tempIndex] = tempString
								if not item_list: item_list = list(item)
								tempString = item[elementIndex]
								tempString += " "
								item_list[elementIndex] = tempString
								itemIndex = elementIndex+3 # skip the chars we just processed
						elif element == '.': # make it pronounce decimals correctly
							if elementIndex == 0 or (elementIndex > 0 and (item[elementIndex-1].isnumeric() or item[elementIndex-1] == ' ')) and elementIndex+1 in range(itemLen) and item[elementIndex+1].isnumeric:
								if not item_list: item_list = list(item)
								item_list[elementIndex] = " point "
								itemIndex = elementIndex+1
								while itemIndex in range(itemLen) and item[itemIndex] == '0':
									item_list[itemIndex] = "o "
									itemIndex+=1
						elif element == ',':
							if elementIndex > 0 and item[elementIndex-1].isnumeric() and elementIndex+1 in range(itemLen) and item[elementIndex+1].isnumeric:
								if not item_list: item_list = list(item)
								item_list[elementIndex] = ""
						elif element == ' ': # stop the synth from saying metric items like liter, gram, etc
							if elementIndex > 0 and item[elementIndex-1].isnumeric():
								itemIndex = elementIndex
								while itemIndex in range(itemLen):
									if not item[itemIndex] == ' ':
										break
									if not item_list: item_list = list(item)
									item_list[itemIndex] = ""
									itemIndex+=1
								if not itemIndex in range(itemLen): # set previousMetric since we are at the end of this string so we can handle it at the beginning of the next string
									self.previousMetric = True
								if itemIndex in range(itemLen) and item[itemIndex].isnumeric(): # don't run numbers previously separated by spaces together
									if not item_list: item_list = list(item)
									item_list[itemIndex-1] = " "
						elif element == '$':
							pointNotFound = True
							if elementIndex > 0 and not item[elementIndex-1] == ' ':
								if not item_list: item_list = list(item)
								item_list[elementIndex] = " $"
							itemIndex = elementIndex+1
							while itemIndex in range(itemLen):
								if not item[itemIndex].isnumeric() and item[itemIndex] != '.' and item[itemIndex] != ',':
									break
								elif item[itemIndex] == ',' and itemIndex+1 in range(itemLen) and item[itemIndex+1].isnumeric:
									if not item_list: item_list = list(item)
									item_list[itemIndex] = ""
								elif item[itemIndex] == '.':
									pointNotFound = False
									if itemIndex+3 in range(itemLen) and item[itemIndex+1].isnumeric() and item[itemIndex+2].isnumeric() and item[itemIndex+3].isnumeric():
										tempIndex = itemIndex+1
										isLeadingZero = True
										moneyString = " and "
										while tempIndex in range(itemLen) and item[tempIndex].isnumeric():
											if not item_list: item_list = list(item)
											item_list[tempIndex] = ""
											if item[tempIndex] == '0' and isLeadingZero:
												moneyString += "zero "
											else:
												isLeadingZero = False
											if item[tempIndex] > '0':
												moneyString += item[tempIndex]
											tempIndex+=1
										moneyString += " cents "
										if not item_list: item_list = list(item)
										item_list[tempIndex-1] = moneyString
										itemIndex = tempIndex-1
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
							if pointNotFound:
								if itemIndex+1 in range(itemLen) and item[itemIndex] == ' ' and item[itemIndex+1].isalpha():
									skipString = "bBmM"
									if item[itemIndex+1] not in skipString:
										if not item_list: item_list = list(item)
										moneyString = ". "
										item_list[itemIndex] = moneyString
										itemIndex+=2
									elif (item[itemIndex+1] == 'b' or item[itemIndex+1] == 'B') and itemIndex+2 in range(itemLen) and (item[itemIndex+2] == 'y' or item[itemIndex+2] == 'Y'):
										if not item_list: item_list = list(item)
										moneyString = ". "
										item_list[itemIndex] = moneyString
										itemIndex+=3
								elif itemIndex in range(itemLen) and item[itemIndex].isalpha():
									skipString = "bBmM"
									if item[itemIndex] not in skipString:
										if not item_list: item_list = list(item)
										moneyString = "."
										moneyString += item[itemIndex]
										item_list[itemIndex] = moneyString
										itemIndex+=1
									elif (item[itemIndex] == 'b' or item[itemIndex] == 'B') and itemIndex+1 in range(itemLen) and (item[itemIndex+1] == 'y' or item[itemIndex+1] == 'Y'):
										if not item_list: item_list = list(item)
										moneyString = "."
										moneyString += item[itemIndex]
										item_list[itemIndex] = moneyString
										itemIndex+=2
						elif element == ':':
							allowOClock = True
							hasSeconds = False
							if elementIndex >= 3 and item[elementIndex-3] == ':':
								hasSeconds = True
							if elementIndex >= 2 and item[elementIndex-1] == '0' and item[elementIndex-2] == '0':
								if not hasSeconds:
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
							if elementIndex == 0 or (elementIndex > 0 and item[elementIndex-1] == ' ' and elementIndex+1 in range(itemLen) and not item[elementIndex+1] == ' '):
								tempIndex = elementIndex
								while tempIndex in range(itemLen):
									if item[tempIndex] == ':':
										tempIndex = 0
										break
									elif item[tempIndex] != '0':
										break
									tempIndex +=1
								if tempIndex:
									itemIndex = tempIndex
									tempIndex -=1
									if not item_list: item_list = list(item)
									while tempIndex >= elementIndex:
										item_list[tempIndex] = "zero "
										tempIndex -=1
				if item_list:
					item = "".join(item_list)
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
				if characterMode and itemLen == 1 and ord(item) in upperAscii:
					item = upperAscii[ord(item)]
				text_list.append(item)
				# when NVDA sends shortcut characters such as alt n it doesn't put a space after the shortcut and this synthesizer needs that to not run it together the next word
				if characterMode:
						text_list.append(" ")
			elif isinstance(item, IndexCommand):
				global lastSentIndex
				global nvdaIndexes
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

		text = "".join(text_list)
		text = text.encode('ascii', 'replace')
		textLength = len(text)
		# only resend the speech parameters when the foreground window changes or when the desktop changes.
		# This accounts for self talking apps that might change the speech parameters and the version of NVDA running on the secure desktop doing the same
		# force this by lying and saying that the variant of the voice has changed because when it changes all parameters have to be resent
		if not api.getForegroundObject() == None:
			if self.lastForegroundProcessID != api.getForegroundObject().processID:
				self.tt_variantChanged = True
				self.lastForegroundProcessID = api.getForegroundObject().processID
		if changedDesktop:
			self.tt_variantChanged = True
			changedDesktop = False
		params = b""
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
		text = b"%s %s %s" % (params, text, b"\r")
		if characterMode or textLength < 10:
			milliseconds = 10 # for short strings use 10 milliseconds to keep things responsive
		else:
			milliseconds = 100 # for long strings use 100 milliseconds as to not hammer the synth for index marks and waste CPU
		# Sometimes the synth can get stuck with the cap pitch offset this forces it to reset pitch after changing pitch offset for caps
		if self.capPitch:
			self.tt_pitchChanged = True
			self.capPitch = False
		# don't use WriteString because it has performance issues, causes other strange behavior, and is just meant for quick testing
		for element in text:
			if element == 0x1e:
				USBTT.USBTT_WriteByteImmediate(element)
			else:
				USBTT.USBTT_WriteByte(element)
		global synthFlushed
		synthFlushed = False
		indexesAvailable.set()

	def cancel(self):
		global synthFlushed
		if synthFlushed or not USBTT:
			return
		synthFlushed = True
		USBTT.USBTT_WriteByte(0x18) # Use WriteByte instead of WriteByteImmediate because the latter can interrupt during command processing causing partial commands to get spoken

	def terminate(self):
		global indexReached
		global stopIndexing
		global synthFlushed
		unload_dll()
		indexReached = None
		stopIndexing = True
		synthFlushed = False
		if not indexesAvailable == None:
			indexesAvailable.set()
		if not self.indexingThread == None:
			self.indexingThread.join()

	def _set_rate(self, rate):
		if rate != self.nvda_rate:
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
		if inflection != self.nvda_inflection:
			self.tt_inflectionChanged = True
			self.tt_inflection = round(inflection/10)
			self.nvda_inflection = inflection
			if self.tt_inflection > self.maxInflection:
				self.tt_inflection = self.maxInflection

	def _get_inflection(self):
		return self.nvda_inflection

	def _set_volume(self, volume):
		if volume != self.nvda_volume:
			self.tt_volumeChanged = True
			self.tt_volume = round(volume/10)
			self.nvda_volume = volume
			if self.tt_volume > self.maxVolume:
				self.tt_volume = self.maxVolume

	def _get_volume(self):
		return self.nvda_volume

	def _getAvailableVariants(self):
		return OrderedDict((str(id), synthDriverHandler.VoiceInfo(str(id), name)) for id, name in self.variants.items())

	def _set_variant(self, v):
		if v != self.tt_variant:
			self.tt_variantChanged = True
			self.tt_variant = v if int(v) in self.variants else "0"

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
			self.pauseModeOn = True
			USBTT.USBTT_WriteByteImmediate(0x10)
		else:
			if self.pauseModeOn:
				self.pauseModeOn = False
				USBTT.USBTT_WriteByteImmediate(0x12)

	def _get_availablePausemodes(self):
		return self.pauseModes

	def _set_pauseMode(self, val):
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
					load_dll(True)
				else:
					global settingPauseMode
					settingPauseMode = True
				self.tt_pauseMode = val

	def _get_pauseMode(self):
		return str(self.tt_pauseMode)
