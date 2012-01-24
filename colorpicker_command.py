import sublime, sublime_plugin
import ctypes
from ctypes import c_int32, c_uint32, c_void_p, c_wchar_p, pointer, POINTER

if sublime.platform() == 'windows':
    class CHOOSECOLOR(ctypes.Structure):
         _fields_ = [('lStructSize', c_uint32),
                     ('hwndOwner', c_void_p),
                     ('hInstance', c_void_p),
                     ('rgbResult', c_uint32),
                     ('lpCustColors',POINTER(c_uint32)),
                     ('Flags', c_uint32),
                     ('lCustData', c_void_p),
                     ('lpfnHook', c_void_p),
                     ('lpTemplateName', c_wchar_p)]

    CustomColorArray = c_uint32 * 16
    CC_SOLIDCOLOR = 0x80
    CC_RGBINIT = 0x01
    CC_FULLOPEN = 0x02

    ChooseColorW = ctypes.windll.Comdlg32.ChooseColorW
    ChooseColorW.argtypes = [POINTER(CHOOSECOLOR)]
    ChooseColorW.restype = c_int32

class PickColorCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        if sublime.platform() != 'windows': 
            sublime.error_message('Sorry, this plugin is for Windows only')
            return

        sel = self.view.sel()
        start_color = None

        if len(sel) > 0:
            # try selecting highlighted text first
            selected = self.view.substr(sel[0]).strip()
            if selected.startswith('#'): selected = selected[1:]
            if self.__is_valid_hex_color(selected):
                start_color = self.__hexstr_to_bgr(selected)
            else:
                # else try selecting by word
                selected = self.view.substr(self.view.word(sel[0])).strip()
                if selected.startswith('#'): selected = selected[1:]
                if self.__is_valid_hex_color(selected):
                    start_color = self.__hexstr_to_bgr(selected)

        s = sublime.load_settings("ColorPicker.sublime-settings")
        custom_colors = s.get("custom_colors", ['0']*16)

        if len(custom_colors) < 16:
            custom_colors = ['0']*16
            s.set('custom_colors', custom_colors)

        cc = CHOOSECOLOR()
        ctypes.memset(ctypes.byref(cc), 0, ctypes.sizeof(cc))
        cc.lStructSize = ctypes.sizeof(cc)
        cc.hwndOwner = self.view.window().hwnd()
        cc.Flags = CC_SOLIDCOLOR | CC_FULLOPEN | CC_RGBINIT
        cc.rgbResult = c_uint32(start_color) if start_color else c_uint32(0)
        cc.lpCustColors = self.__to_custom_color_array(custom_colors)

        if ChooseColorW(ctypes.byref(cc)):
            color = self.__bgr_to_hexstr(cc.rgbResult)
            for region in sel:
                word = self.view.word(region)
                if self.__is_valid_hex_color(self.view.substr(word)):
                    if self.view.substr(word.a - 1) == '#':
                        word = sublime.Region(word.a - 1, word.b)
                    self.view.replace(edit, word, '#' + color)
                else:
                    self.view.replace(edit, region, '#' + color)

        custom_colors = self.__from_custom_color_array(cc.lpCustColors)
        s.set('custom_colors', custom_colors)
        sublime.save_settings("ColorPicker.sublime-settings")

    def __to_custom_color_array(self, custom_colors):
        cc = CustomColorArray()
        for i in range(16):
            cc[i] = int(custom_colors[i])
        return cc

    def __from_custom_color_array(self, custom_colors):
        cc = [0]*16
        for i in range(16):
            cc[i] = str(custom_colors[i])
        return cc

    def __is_valid_hex_color(self, s):
        if len(s) not in (3, 6):
            return False
        try:
            return 0 <= int(s, 16) <= 0xffffff
        except ValueError:
            return False

    def __bgr_to_hexstr(self, bgr, byte_table=list(map(lambda b: '{0:02X}'.format(b), range(256)))):
        # 0x00BBGGRR
        b = byte_table[(bgr >> 16) & 0xff]
        g = byte_table[(bgr >>  8) & 0xff]
        r = byte_table[(bgr      ) & 0xff]
        return (r+g+b)

    def __hexstr_to_bgr(self, hexstr):
        if len(hexstr) == 3:
            hexstr = hexstr[0] + hexstr[0] + hexstr[1] + hexstr[1] + hexstr[2] + hexstr[2]
            
        r = int(hexstr[0:2], 16)
        g = int(hexstr[2:4], 16)
        b = int(hexstr[4:6], 16)
        return (b << 16)| (g << 8) | r
