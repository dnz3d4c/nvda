#A part of NonVisual Desktop Access (NVDA)
#Copyright (C) 2006-2017 NV Access Limited, Manish Agrawal, Derek Riemer, Babbage B.V.
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

import ui
import textInfos
from . import Window

class WordDocument(Window):
	"""
	An abstract class for WordDocuments which only contains script bindings and abstract code common to both Object Model and UI automation implementations for MS Word
	It uses the same original name for the WordDocument class ensuring that existing key bindings are not broken.
	"""

	def script_moveParagraphDown(self,gesture):
		oldBookmark=self.makeTextInfo(textInfos.POSITION_CARET).bookmark
		gesture.send()
		if self._hasCaretMoved(oldBookmark)[0]:
			info=self.makeTextInfo(textInfos.POSITION_SELECTION)
			info.collapse()
			info.move(textInfos.UNIT_PARAGRAPH,-1,endPoint="start")
			lastParaText=info.text.strip()
			if lastParaText:
				# Translators: a message reported when a paragraph is moved below another paragraph
				ui.message(_("Moved below %s")%lastParaText)
			else:
				# Translators: a message reported when a paragraph is moved below a blank paragraph 
				ui.message(_("Moved below blank paragraph"))

	def script_moveParagraphUp(self,gesture):
		oldBookmark=self.makeTextInfo(textInfos.POSITION_CARET).bookmark
		gesture.send()
		if self._hasCaretMoved(oldBookmark)[0]:
			info=self.makeTextInfo(textInfos.POSITION_SELECTION)
			info.collapse()
			info.move(textInfos.UNIT_PARAGRAPH,1)
			info.expand(textInfos.UNIT_PARAGRAPH)
			lastParaText=info.text.strip()
			if lastParaText:
				# Translators: a message reported when a paragraph is moved above another paragraph
				ui.message(_("Moved above %s")%lastParaText)
			else:
				# Translators: a message reported when a paragraph is moved above a blank paragraph 
				ui.message(_("Moved above blank paragraph"))

	__gestures = {
		"kb:control+[":"increaseDecreaseFontSize",
		"kb:control+]":"increaseDecreaseFontSize",
		"kb:control+shift+,":"increaseDecreaseFontSize",
		"kb:control+shift+.":"increaseDecreaseFontSize",
		"kb:control+b":"toggleBold",
		"kb:control+i":"toggleItalic",
		"kb:control+u":"toggleUnderline",
		"kb:control+=":"toggleSuperscriptSubscript",
		"kb:control+shift+=":"toggleSuperscriptSubscript",
		"kb:control+l":"toggleAlignment",
		"kb:control+e":"toggleAlignment",
		"kb:control+r":"toggleAlignment",
		"kb:control+j":"toggleAlignment",
		"kb:alt+shift+downArrow":"moveParagraphDown",
		"kb:alt+shift+upArrow":"moveParagraphUp",
		"kb:alt+shift+rightArrow":"increaseDecreaseOutlineLevel",
		"kb:alt+shift+leftArrow":"increaseDecreaseOutlineLevel",
		"kb:control+shift+n":"increaseDecreaseOutlineLevel",
		"kb:control+alt+1":"increaseDecreaseOutlineLevel",
		"kb:control+alt+2":"increaseDecreaseOutlineLevel",
		"kb:control+alt+3":"increaseDecreaseOutlineLevel",
		"kb:control+1":"changeLineSpacing",
		"kb:control+2":"changeLineSpacing",
		"kb:control+5":"changeLineSpacing",
		"kb:tab": "tab",
		"kb:shift+tab": "tab",
		"kb:alt+home":"caret_moveByCell",
		"kb:alt+end":"caret_moveByCell",
		"kb:alt+pageUp":"caret_moveByCell",
		"kb:alt+pageDown":"caret_moveByCell",
		"kb:alt+shift+home":"caret_changeSelection",
		"kb:alt+shift+end":"caret_changeSelection",
		"kb:alt+shift+pageUp":"caret_changeSelection",
		"kb:alt+shift+pageDown":"caret_changeSelection",
		"kb:control+pageUp": "caret_moveByLine",
		"kb:control+pageDown": "caret_moveByLine",
		"kb:NVDA+alt+c":"reportCurrentComment",
	}
