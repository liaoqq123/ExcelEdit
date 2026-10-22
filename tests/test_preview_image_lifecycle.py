"""无需显示环境的图片引用回归测试；不创建 Tk 或任何窗口。"""

from types import SimpleNamespace
from tkinter import TclError
import unittest
import weakref

from result_preview import ResultPreviewPane, _set_label_image


class FakeImage:
    pass


class FakeNativeLabel:
    """模拟 Tcl 仅保存图片名，不能延长 Python PhotoImage 的寿命。"""

    def __init__(self) -> None:
        self.image_reference = None
        self.text = ""

    def configure(self, **kwargs) -> None:
        if "image" in kwargs:
            image = kwargs["image"]
            self.image_reference = weakref.ref(image) if image else None
        if self.image_reference is not None and self.image_reference() is None:
            raise TclError('image "pyimage3" doesn\'t exist')
        if "text" in kwargs:
            self.text = kwargs["text"]


class FakeCTkLabel:
    """保留 CTkLabel 6.0 的 text-first / image=None 不清原生图片行为。"""

    def __init__(self, image: FakeImage) -> None:
        self._label = FakeNativeLabel()
        self._image = image
        self._label.configure(image=image)

    def configure(self, **kwargs) -> None:
        if "text" in kwargs:
            self._label.configure(text=kwargs["text"])
        if "image" in kwargs:
            self._image = kwargs["image"]
            if self._image is not None:
                self._label.configure(image=self._image)
        self._label.configure()


class PreviewImageLifecycleTests(unittest.TestCase):
    def test_old_combined_configure_reproduces_missing_photoimage(self) -> None:
        label = FakeCTkLabel(FakeImage())
        with self.assertRaisesRegex(TclError, "pyimage3"):
            label.configure(image=None, text="")

    def test_detach_clears_native_image_before_ctk_releases_reference(self) -> None:
        label = FakeCTkLabel(FakeImage())
        previous = weakref.ref(label._image)
        _set_label_image(label, None, "正在准备预览…")
        self.assertIsNone(previous())
        self.assertIsNone(label._label.image_reference)
        self.assertIsNone(label._image)
        self.assertEqual(label._label.text, "正在准备预览…")
        label.configure(text="预览暂不可用")

    def test_replacement_retains_new_image_and_releases_previous_image(self) -> None:
        label = FakeCTkLabel(FakeImage())
        previous = weakref.ref(label._image)
        replacement = FakeImage()
        _set_label_image(label, replacement, "")
        self.assertIsNone(previous())
        self.assertIs(label._label.image_reference(), replacement)
        self.assertIs(label._image, replacement)
        label.configure(text="安全更新")

    def test_detach_can_recover_an_already_stale_native_reference(self) -> None:
        label = FakeCTkLabel(FakeImage())
        label._image = None
        with self.assertRaises(TclError):
            label._label.configure(text="旧引用仍存在")
        _set_label_image(label, None, "已清除")
        self.assertIsNone(label._label.image_reference)
        self.assertEqual(label._label.text, "已清除")

    def test_pane_reference_is_released_after_native_detachment(self) -> None:
        label = FakeCTkLabel(FakeImage())
        previous = weakref.ref(label._image)
        pane = SimpleNamespace(file_image_label=label, _image=label._image)
        ResultPreviewPane._set_file_image(pane, None, "")
        self.assertIsNone(previous())
        self.assertIsNone(pane._image)
        self.assertIsNone(label._label.image_reference)


if __name__ == "__main__":
    unittest.main()
