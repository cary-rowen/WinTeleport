"""WinTeleport - Move windows between virtual desktops.

Provides keyboard commands to move the focused window between Windows 10/11 virtual desktops.
"""

import functools
import os
import sys
from collections.abc import Callable
from typing import TypeVar

import api
import addonHandler
import globalPluginHandler
import ui
import winUser
from comtypes import CLSCTX_ALL, COMError, CoCreateInstance
from inputCore import InputGesture
from logHandler import log
from scriptHandler import script

addonHandler.initTranslation()

# Add pyvda submodule to sys.path
_ADDON_PATH = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PYVDA_PATH = os.path.join(_ADDON_PATH, "pyvdaRepo")
if _PYVDA_PATH not in sys.path:
	sys.path.insert(0, _PYVDA_PATH)

from pyvda import AppView, VirtualDesktop, get_virtual_desktops
from pyvda.com_defns import CLSID_IVirtualDesktopManager, IVirtualDesktopManager
import pyvda.utils


def _hresultFromWin32(code: int) -> int:
	"""Convert a Win32 error code to HRESULT."""
	return (0x80070000 | code) - 0x100000000


# RPC_S_SERVER_UNAVAILABLE - COM connection lost
_HRESULT_RPC_UNAVAILABLE = _hresultFromWin32(1722)
# TYPE_E_ELEMENTNOTFOUND - Window has no IApplicationView
_HRESULT_ELEMENT_NOT_FOUND = -2147319765


def _isRpcUnavailableError(error: COMError) -> bool:
	return error.args[0] == _HRESULT_RPC_UNAVAILABLE


def _isElementNotFoundError(error: COMError) -> bool:
	return error.args[0] == _HRESULT_ELEMENT_NOT_FOUND


def _getForegroundWindowHandle() -> int:
	"""Get the root owner window handle of the current foreground window."""
	fg = api.getForegroundObject()
	if fg is None:
		return 0
	hwnd = getattr(fg, "windowHandle", 0)
	if not hwnd:
		return 0
	rootHwnd = winUser.getAncestor(hwnd, winUser.GA_ROOTOWNER)
	return rootHwnd or hwnd


def _getDesktopDisplayName(desktop: VirtualDesktop) -> str:
	"""Get a display name for a virtual desktop."""
	try:
		if desktop.name:
			return desktop.name
	except (NotImplementedError, AttributeError, OSError):
		pass
	# Translators: Default name for a virtual desktop. {number} is the desktop number.
	return _("Desktop {number}").format(number=desktop.number)


_F = TypeVar("_F", bound=Callable)


def _withComRetry(func: _F) -> _F:
	"""Decorator that retries COM operations once if RPC becomes unavailable.

	When Windows Shell restarts, COM connections become stale. This decorator
	catches RPC_S_SERVER_UNAVAILABLE errors, reinitializes the COM managers,
	and retries the operation once silently.
	"""

	@functools.wraps(func)
	def wrapper(self, *args, **kwargs):
		for attempt in range(2):
			try:
				return func(self, *args, **kwargs)
			except COMError as e:
				if attempt == 0 and _isRpcUnavailableError(e):
					log.debugWarning("RPC unavailable, reinitializing pyvda and retrying")
					self._reinitializePyvda()
					continue
				log.debugWarning("COM error in %s: %s", func.__name__, e)
				raise

	return wrapper  # type: ignore


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	# Translators: The name of the category for this add-on's scripts in the Input Gestures dialog.
	scriptCategory = _("WinTeleport")

	def __init__(self):
		super().__init__()
		self._vdManager: IVirtualDesktopManager | None = None

	def _reinitializePyvda(self) -> None:
		"""Reinitialize pyvda's COM managers after connection loss."""
		log.debug("Reinitializing pyvda COM managers")
		newManagers = pyvda.utils.Managers()
		pyvda.utils.managers = newManagers
		# Update the reference in pyvda.pyvda module as well
		if hasattr(pyvda, "pyvda") and hasattr(pyvda.pyvda, "managers"):
			pyvda.pyvda.managers = newManagers
		self._vdManager = None

	def _getVirtualDesktopManager(self) -> IVirtualDesktopManager:
		"""Get or create a cached IVirtualDesktopManager instance."""
		if self._vdManager is None:
			self._vdManager = CoCreateInstance(
				CLSID_IVirtualDesktopManager,
				IVirtualDesktopManager,
				CLSCTX_ALL,
			)
		return self._vdManager

	def _moveWindowToDesktop(self, hwnd: int, targetDesktop: VirtualDesktop) -> bool:
		"""Move a window to a target desktop.

		Tries AppView.move() first for full feature support,
		falls back to IVirtualDesktopManager for broader compatibility.
		"""
		# Try AppView first (supports pinned apps, etc.)
		try:
			AppView(hwnd=hwnd).move(targetDesktop)
			return True
		except COMError as e:
			if not _isElementNotFoundError(e):
				log.error("COM error moving window via AppView: %s", e)
				return False
			log.debug("AppView not available for hwnd %#x, trying fallback", hwnd)
		# Fallback to IVirtualDesktopManager
		try:
			manager = self._getVirtualDesktopManager()
			manager.MoveWindowToDesktop(hwnd, targetDesktop.id)
			return True
		except COMError as e:
			log.error("COM error moving window via VDManager: %s", e)
			return False

	def _focusNextWindow(self) -> None:
		"""Focus the next visible window on the current desktop.

		Called after moving a window away to prevent focus from being lost.
		"""
		try:
			currentDesktop = VirtualDesktop.current()
			apps = currentDesktop.apps_by_z_order()
			if apps:
				apps[0].set_focus()
		except COMError:
			pass  # Silently fail if no window to focus

	@script(
		# Translators: Describes a command in Input Help mode and Input Gestures dialog.
		description=_("Move the focused window to the left virtual desktop"),
	)
	def script_moveWindowToLeftDesktop(self, gesture: InputGesture) -> None:
		self._moveToAdjacentDesktop(direction=-1, followWindow=False)

	@script(
		# Translators: Describes a command in Input Help mode and Input Gestures dialog.
		description=_("Move the focused window to the right virtual desktop"),
	)
	def script_moveWindowToRightDesktop(self, gesture: InputGesture) -> None:
		self._moveToAdjacentDesktop(direction=1, followWindow=False)

	@script(
		# Translators: Describes a command in Input Help mode and Input Gestures dialog.
		description=_("Move the focused window to the left virtual desktop and switch to it"),
		gesture="kb:NVDA+control+windows+leftArrow",
	)
	def script_moveWindowToLeftDesktopAndFollow(self, gesture: InputGesture) -> None:
		self._moveToAdjacentDesktop(direction=-1, followWindow=True)

	@script(
		# Translators: Describes a command in Input Help mode and Input Gestures dialog.
		description=_("Move the focused window to the right virtual desktop and switch to it"),
		gesture="kb:NVDA+control+windows+rightArrow",
	)
	def script_moveWindowToRightDesktopAndFollow(self, gesture: InputGesture) -> None:
		self._moveToAdjacentDesktop(direction=1, followWindow=True)

	@script(
		# Translators: Describes a command in Input Help mode and Input Gestures dialog.
		description=_("Report the current virtual desktop"),
		gesture="kb:NVDA+windows+d",
	)
	def script_reportCurrentDesktop(self, gesture: InputGesture) -> None:
		self._reportCurrentDesktop()

	@_withComRetry
	def _reportCurrentDesktop(self) -> None:
		"""Report the current virtual desktop name and number."""
		currentDesktop = VirtualDesktop.current()
		desktopCount = len(get_virtual_desktops())
		name = _getDesktopDisplayName(currentDesktop)
		# Translators: Message for reporting current desktop.
		ui.message(
			_("{name} ({number} of {total})").format(
				name=name,
				number=currentDesktop.number,
				total=desktopCount,
			),
		)

	def script_moveWindowToDesktopN(self, gesture: InputGesture) -> None:
		try:
			number = int(gesture.mainKeyName)
		except ValueError:
			log.error("Invalid key name for desktop number: %s", gesture.mainKeyName)
			return
		self._moveToDesktopNumber(number, followWindow=False)

	# Translators: Describes a command in Input Help mode and Input Gestures dialog.
	script_moveWindowToDesktopN.__doc__ = _("Move the focused window to the specified virtual desktop")

	@_withComRetry
	def _moveToAdjacentDesktop(self, direction: int, followWindow: bool) -> None:
		"""Move the focused window to an adjacent desktop."""
		hwnd = _getForegroundWindowHandle()
		if not hwnd:
			# Translators: Error message when no window is focused.
			ui.message(_("No focused window to move"))
			return
		desktops = get_virtual_desktops()
		if len(desktops) < 2:
			# Translators: Message when there is only one virtual desktop.
			ui.message(_("Only one desktop available"))
			return
		currentDesktop = VirtualDesktop.current()
		targetIndex = currentDesktop.number - 1 + direction
		if targetIndex < 0:
			# Translators: Message when already on the first desktop.
			ui.message(_("Already on the first desktop"))
			return
		if targetIndex >= len(desktops):
			# Translators: Message when already on the last desktop.
			ui.message(_("Already on the last desktop"))
			return
		targetDesktop = desktops[targetIndex]
		if not self._moveWindowToDesktop(hwnd, targetDesktop):
			# Translators: Error message when window move fails.
			ui.message(_("Failed to move window"))
			return
		if followWindow:
			targetDesktop.go()
		else:
			self._focusNextWindow()

	@_withComRetry
	def _moveToDesktopNumber(self, number: int, followWindow: bool) -> None:
		"""Move the focused window to a specific desktop number."""
		hwnd = _getForegroundWindowHandle()
		if not hwnd:
			# Translators: Error message when no window is focused.
			ui.message(_("No focused window to move"))
			return
		desktops = get_virtual_desktops()
		if number > len(desktops):
			# Translators: Message when desktop doesn't exist.
			ui.message(
				_("Desktop {number} does not exist. {total} desktops available.").format(
					number=number,
					total=len(desktops),
				),
			)
			return
		targetDesktop = desktops[number - 1]
		if not self._moveWindowToDesktop(hwnd, targetDesktop):
			# Translators: Error message when window move fails.
			ui.message(_("Failed to move window"))
			return
		if not followWindow:
			self._focusNextWindow()
		self._announceMove(targetDesktop, followWindow)

	def _announceMove(self, targetDesktop: VirtualDesktop, followWindow: bool) -> None:
		"""Announce the result of a window move operation."""
		targetName = _getDesktopDisplayName(targetDesktop)
		if followWindow:
			targetDesktop.go()
			# Translators: Message after moving window and switching desktop.
			ui.message(_("Moved and switched to {name}").format(name=targetName))
		else:
			# Translators: Message after moving window to another desktop.
			ui.message(_("Moved to {name}").format(name=targetName))

	__gestures = {f"kb:NVDA+control+windows+{n}": "moveWindowToDesktopN" for n in range(1, 10)}
