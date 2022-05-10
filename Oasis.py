"""
패치노트 1.0.1
1. 파일명 변경시 한국어를 무시하는 오류 수정
2. 라이브러리 openpyxl 추가
3. 일부 어색한 문구 수정
4. 시트 저장 관련 오류수정 및 최적화
패치노트 1.0.2
1. 딥러닝 장치 CPU 옵션 추가
2. GPU 호환이 되지 않을 경우에도 자동으로 CPU를 사용하도록 적용
3. 프로그램 사용 도중 사진을 읽어올 수 없을 경우 더미파일이 생성되는 문제 수정
패치노트 1.0.3
1. 임시저장이 되지 않는 오류 수정
2. 프로그램 로그를 기록하는 기능을 추가
3. 프로그램이 비정상종료 되는 치명적인 오류 발생 시 종료 대신 메세지를 띄우도록 수정
패치노트 1.0.4
1. 단축키가 간헐적으로 작동하지 않는 문제 완화
 (작동하지 않을 경우 esc 키를 한번씩 눌러주세요.)
2. 동물종 목록 시스템 개편
 - 드래그&드롭으로 동물종 명의 위치를 수정 가능
 - 기본제공 동물 DB 추가
  (포유류 125종, 조류 545종, 파충류32종 합 702종)
 - 국명으로 추가시 학명 자동 등록
"""
import sys
import os
import re
import time
import platform
import traceback
import win32file
import yaml
import math
import shutil
import warnings
import logging
import requests
from copy import deepcopy
from glob import glob
from pathlib import Path
from datetime import datetime, timedelta
from collections import Counter
from PIL import ImageFont, ImageDraw, Image
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtTest import QTest
from natsort import natsorted
import numpy as np
import pandas as pd
import pkg_resources as pkg
import cv2
import torch
import torch.nn as nn
import torchvision
# pyinstaller에 묶기 위한 trash import ###
import seaborn
import scipy
import utils
import models
import openpyxl

##########################################

version = '1.0.4'

app = QApplication(sys.argv)
screen = app.primaryScreen()
size = screen.size()
width = size.width()
height = size.height()
FILE = Path(__file__).absolute()
RANK = int(os.getenv('RANK', -1))
sys.path.append(FILE.parents[0].as_posix())  # add yolov5/ to path

IMG_FORMATS = ['bmp', 'jpg', 'jpeg', 'png', 'tif', 'tiff', 'dng', 'webp', 'mpo']  # acceptable image suffixes
VID_FORMATS = ['mov', 'avi', 'mp4', 'mpg', 'mpeg', 'm4v', 'wmv', 'mkv']  # acceptable video suffixes
DEFAULT_COLUMNS = ['파일명', '년', '월', '일', '시', '분', '주야간', '국명', '학명', '개체수', '온도', '최대정확도', '최소정확도']  # 고정 칼럼
DEFAULT_ANIMAL = ['미동정', '멧돼지', '고라니', '산양']
icon = QIcon('./icon/logo2.png')


def check_update(version='1.0.0', test=False):
    if test:
        page = requests.get(f'https://tripleler.tistory.com/entry/Oasis-100')
        if page.status_code == 200:
            print('update check!')
            return
    else:
        next_version = [str(int(version.split('.')[0]) + 1) + '.0.0',
                        version.split('.')[0] + '.' + str(int(version.split('.')[1]) + 1) + '.0',
                        version.split('.')[0] + '.' + version.split('.')[1] + '.' + str(int(version.split('.')[2]) + 1)]
        for v in next_version:
            page = requests.get(f'https://tripleler.tistory.com/entry/Oasis-{v}')
            if page.status_code == 200:
                return True
        return False


def time_sync():
    # pytorch-accurate time
    if torch.cuda.is_available():
        torch.cuda.synchronize()
    return time.time()


# APPDATA 접근 =========================================================================================================

def is_writeable(dir):  # utils.general
    # Return True if directory has write permissions, test opening a file with write permissions
    status = os.access(dir, os.F_OK & os.R_OK & os.W_OK & os.X_OK)
    if status:
        file = Path(dir) / 'tmp.txt'
        try:
            with open(file, 'w'):  # open file with write permissions
                pass
            file.unlink()  # remove file
            return True
        except OSError:
            return False
    else:
        return False


def user_config_dir(dir='Oasis'):  # utils.general
    # Return path of user configuration directory. Prefer environment variable if exists. Make dir if required.
    cfg = {'Windows': 'AppData/Roaming', 'Linux': '.config', 'Darwin': 'Library/Application Support'}  # 3 OS dirs
    path = Path.home() / cfg.get(platform.system(), '')  # OS-specific config dir
    path = (path if is_writeable(path) else Path('/tmp')) / dir  # GCP and AWS lambda fix, only /tmp is writeable
    path.mkdir(exist_ok=True, parents=True)  # make if required
    log = path / 'logs'  # save log dir
    log.mkdir(exist_ok=True, parents=True)
    return path


CONFIG_DIR = user_config_dir()


# End APPDATA =========================================================================================================
# Check data ==========================================================================================================

def check_font(font='Arial.ttf', size=10):
    # Return a PIL TrueType Font, downloading to CONFIG_DIR if necessary
    font = Path(font)
    font = font if font.exists() else (CONFIG_DIR / font.name)
    try:
        return ImageFont.truetype(str(font) if font.exists() else font.name, size)
    except Exception:  # download if missing
        url = "https://ultralytics.com/assets/" + font.name
        print(f'Downloading {url} to {font}...')
        torch.hub.download_url_to_file(url, str(font), progress=False)
        return ImageFont.truetype(str(font), size)


def check_appdata(file='app_data.yaml'):
    dir = CONFIG_DIR / file
    if not dir.exists():
        return {'columns': DEFAULT_COLUMNS, 'category': DEFAULT_ANIMAL}
    else:
        with open(dir, 'r', encoding='utf-8') as f:
            data = yaml.load(f, Loader=yaml.FullLoader)
        if 'columns' not in data.keys() or not isinstance(data['columns'], list) \
                or data['columns'][:len(DEFAULT_COLUMNS)] != DEFAULT_COLUMNS:
            data['columns'] = DEFAULT_COLUMNS
        if 'category' not in data.keys() or not isinstance(data['category'], list) \
                or len([x for x in DEFAULT_ANIMAL if x in data['category']]) != len(DEFAULT_ANIMAL):
            data['category'] = DEFAULT_ANIMAL
        return data


def check_animal_db(file='animal_db.yaml'):
    animal = CONFIG_DIR / file
    if not animal.exists():
        try:
            shutil.copy('animal_db.yaml', CONFIG_DIR / file)
        except FileNotFoundError:
            logger.warning('Cannot find "animal_db.yaml"')
            return {}
    with open(animal, 'r', encoding='utf-8') as f:
        animal = yaml.load(f, Loader=yaml.FullLoader)
    return animal


def check_cache(file='cache.csv'):
    cache = CONFIG_DIR / file
    if cache.exists():
        try:
            cache_in = pd.read_csv(cache, keep_default_na=False)
        except Exception as e:
            logger.error('check_cache error', exc_info=e)
        else:
            return cache_in


# End Check Data ======================================================================================================
# logger & exception ==================================================================================================

logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

ch = logging.StreamHandler()  # print console

today = datetime.now().strftime('%Y%m%d')
logfile = CONFIG_DIR / 'logs' / f"{today}.log"
fh = logging.FileHandler(filename=logfile)  # logging file
fh.setLevel(logging.WARNING)  # only logging over warning level
fh.setFormatter(logging.Formatter("%(levelname)s %(asctime)s - %(message)s"))

logger.addHandler(ch)  # add handlers to logger
logger.addHandler(fh)
logger.warning(f'APP START! torch version:{torch.__version__}\nGPU:{torch.cuda.is_available()}')


def show_exception_box(log_msg):
    """Checks if a QApplication instance is available and shows a messagebox with the exception message.
    If unavailable (non-console application), log an additional notice.
    """
    if QApplication.instance() is not None:
        errorbox = QMessageBox()
        errorbox.setWindowIcon(icon)
        errorbox.setWindowTitle('Error')
        errorbox.setText(f"예상치 못한 에러가 발생하였습니다.\n현재 작업내용을 저장하고, 오류를 보고하세요.\n오류내역보고서 저장위치:\n{logfile}")
        errorbox.exec_()


class UncaughtHook(QObject):
    _exception_caught = pyqtSignal(object)

    def __init__(self, *args, **kwargs):
        super(UncaughtHook, self).__init__(*args, **kwargs)

        # this registers the exception_hook() function as hook with the Python interpreter
        sys.excepthook = self.exception_hook

        # connect signal to execute the message box function always on main thread
        self._exception_caught.connect(show_exception_box)

    def exception_hook(self, exc_type, exc_value, exc_traceback):
        """Function handling uncaught exceptions.
        It is triggered each time an uncaught exception occurs.
        """
        if issubclass(exc_type, KeyboardInterrupt):
            # ignore traceback to auto close app. except keyboard interrupt
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
        else:
            exc_info = (exc_type, exc_value, exc_traceback)
            log_msg = '\n'.join([''.join(traceback.format_tb(exc_traceback)),
                                 f'{exc_type.__name__}: {exc_value}'])
            logger.critical(f"Uncaught exception:\n {log_msg}", exc_info=exc_info)

            # trigger message box show
            self._exception_caught.emit(log_msg)


qt_exception_hook = UncaughtHook()


# End logger & Exception ==============================================================================================


def locate_usb():
    drive_list = []
    drivebits = win32file.GetLogicalDrives()
    for d in range(1, 26):
        mask = 1 << d
        if drivebits & mask:
            # here if the drive is at least there
            drname = '%c:\\' % chr(ord('A') + d)
            t = win32file.GetDriveType(drname)
            if t == win32file.DRIVE_REMOVABLE:
                drive_list.append(drname[:2])
    return drive_list


DRIVE = locate_usb()


def is_ascii(s=''):
    # Is string composed of all ASCII (no UTF) characters? (note str().isascii() introduced in python 3.7)
    s = str(s)  # convert list, tuple, None, etc. to str
    return len(s.encode().decode('ascii', 'ignore')) == len(s)


def autopad(k, p=None):  # kernel, padding models.common
    # Pad to 'same' /
    if p is None:
        p = k // 2 if isinstance(k, int) else [x // 2 for x in k]  # auto-pad
    return p


class Conv(nn.Module):  # models.common
    # Standard convolution
    def __init__(self, c1, c2, k=1, s=1, p=None, g=1, act=True):  # ch_in, ch_out, kernel, stride, padding, groups
        super().__init__()
        self.conv = nn.Conv2d(c1, c2, k, s, autopad(k, p), groups=g, bias=False)
        self.bn = nn.BatchNorm2d(c2)
        self.act = nn.SiLU() if act is True else (act if isinstance(act, nn.Module) else nn.Identity())

    def forward(self, x):
        return self.act(self.bn(self.conv(x)))

    def forward_fuse(self, x):
        return self.act(self.conv(x))


class Ensemble(nn.ModuleList):  # models.experimental
    # Ensemble of models
    def __init__(self):
        super().__init__()

    def forward(self, x, augment=False, profile=False, visualize=False):
        y = []
        for module in self:
            y.append(module(x, augment, profile, visualize)[0])
        # y = torch.stack(y).max(0)[0]  # max ensemble
        # y = torch.stack(y).mean(0)  # mean ensemble
        y = torch.cat(y, 1)  # nms ensemble
        return y, None  # inference, train output


def attempt_load(weights, map_location=None, inplace=True, fuse=True, info=None):  # models.experimental
    from models.yolo import Detect, Model
    # Loads an ensemble of models weights=[a,b,c] or a single model weights=[a] or weights=a
    model = Ensemble()
    for w in weights if isinstance(weights, list) else [weights]:
        ckpt = torch.load(w, map_location=map_location)  # load
        if fuse:
            try:
                model.append(ckpt['ema' if ckpt.get('ema') else 'model'].float().fuse().eval())  # FP32 model
            except RuntimeError:
                map_location = torch.device('cpu')
                info = '   GPU 사용불가     '
                ckpt = torch.load(w, map_location=map_location)  # load
                model.append(ckpt['ema' if ckpt.get('ema') else 'model'].float().fuse().eval())  # FP32 model
        else:
            model.append(ckpt['ema' if ckpt.get('ema') else 'model'].float().eval())  # without layer fuse

    # Compatibility updates
    for m in model.modules():
        if type(m) in [nn.Hardswish, nn.LeakyReLU, nn.ReLU, nn.ReLU6, nn.SiLU, Detect, Model]:
            m.inplace = inplace  # pytorch 1.7.0 compatibility
            if type(m) is Detect:
                if not isinstance(m.anchor_grid, list):  # new Detect Layer compatibility
                    delattr(m, 'anchor_grid')
                    setattr(m, 'anchor_grid', [torch.zeros(1)] * m.nl)
        elif type(m) is Conv:
            m._non_persistent_buffers_set = set()  # pytorch 1.6.0 compatibility

    if len(model) == 1:
        return model[-1], map_location, info  # return model
    else:
        print(f'Ensemble created with {weights}\n')
        for k in ['names']:
            setattr(model, k, getattr(model[-1], k))
        model.stride = model[torch.argmax(torch.tensor([m.stride.max() for m in model])).int()].stride  # max stride
        return model, map_location, info  # return ensemble


def increment_path(path, exist_ok=False, sep='', mkdir=False):  # utils.general
    # Increment file or directory path, i.e. runs/exp --> runs/exp{sep}2, runs/exp{sep}3, ... etc.
    path = Path(path)  # os-agnostic
    if path.exists() and not exist_ok:
        suffix = path.suffix
        path = path.with_suffix('')
        dirs = glob(f"{path}{sep}*")  # similar paths
        matches = [re.search(rf"%s{sep}(\d+)" % path.stem, d) for d in dirs]
        i = [int(m.groups()[0]) for m in matches if m]  # indices
        n = max(i) + 1 if i else 2  # increment number
        path = Path(f"{path}{sep}{n}{suffix}")  # update path
    dir = path if path.suffix == '' else path.parent  # directory
    if not dir.exists() and mkdir:
        dir.mkdir(parents=True, exist_ok=True)  # make directory
    return path


def select_device(device='', batch_size=None):
    # device = 'cpu' or '0' or '0,1,2,3' / utils.torch_utils
    s = f'YOLOv5 torch {torch.__version__} '  # string
    device = str(device).strip().lower().replace('cuda:', '')  # to string, 'cuda:0' to '0'
    cpu = device == 'cpu'
    if cpu:
        os.environ['CUDA_VISIBLE_DEVICES'] = '-1'  # force torch.cuda.is_available() = False
        info = '   CPU 사용중     '
    elif device:  # non-cpu device requested
        os.environ['CUDA_VISIBLE_DEVICES'] = device  # set environment variable
        assert torch.cuda.is_available(), f'CUDA unavailable, invalid device {device} requested'  # check availability

    cuda = not cpu and torch.cuda.is_available()
    if cuda:
        devices = device.split(',') if device else '0'  # range(torch.cuda.device_count())  # i.e. 0,1,6,7
        n = len(devices)  # device count
        if n > 1 and batch_size:  # check batch_size is divisible by device_count
            assert batch_size % n == 0, f'batch-size {batch_size} not multiple of GPU count {n}'
        space = ' ' * (len(s) + 1)
        for i, d in enumerate(devices):
            p = torch.cuda.get_device_properties(i)
            s += f"{'' if i == 0 else space}CUDA:{d} ({p.name}, {p.total_memory / 1024 ** 2}MB)\n"  # bytes to MB
            info = f'   GPU 사용중 ({p.name}, {p.total_memory / 1024 ** 2}MB)     '
    else:
        s += 'CPU\n'
    # print(s)  # emoji-safe
    return torch.device('cuda:0' if cuda else 'cpu'), info


def letterbox(im, new_shape=(640, 640), color=(114, 114, 114), auto=True, scaleFill=False, scaleup=True, stride=32):
    # utils.augmentations
    # Resize and pad image while meeting stride-multiple constraints
    shape = im.shape[:2]  # current shape [height, width]
    if isinstance(new_shape, int):
        new_shape = (new_shape, new_shape)

    # Scale ratio (new / old)
    r = min(new_shape[0] / shape[0], new_shape[1] / shape[1])
    if not scaleup:  # only scale down, do not scale up (for better val mAP)
        r = min(r, 1.0)

    # Compute padding
    ratio = r, r  # width, height ratios
    new_unpad = int(round(shape[1] * r)), int(round(shape[0] * r))
    dw, dh = new_shape[1] - new_unpad[0], new_shape[0] - new_unpad[1]  # wh padding
    if auto:  # minimum rectangle
        dw, dh = np.mod(dw, stride), np.mod(dh, stride)  # wh padding
    elif scaleFill:  # stretch
        dw, dh = 0.0, 0.0
        new_unpad = (new_shape[1], new_shape[0])
        ratio = new_shape[1] / shape[1], new_shape[0] / shape[0]  # width, height ratios

    dw /= 2  # divide padding into 2 sides
    dh /= 2

    if shape[::-1] != new_unpad:  # resize
        im = cv2.resize(im, new_unpad, interpolation=cv2.INTER_LINEAR)
    top, bottom = int(round(dh - 0.1)), int(round(dh + 0.1))
    left, right = int(round(dw - 0.1)), int(round(dw + 0.1))
    im = cv2.copyMakeBorder(im, top, bottom, left, right, cv2.BORDER_CONSTANT, value=color)  # add border
    return im, ratio, (dw, dh)


def xywh2xyxy(x):  # utils.general
    # Convert nx4 boxes from [x, y, w, h] to [x1, y1, x2, y2] where xy1=top-left, xy2=bottom-right
    y = x.clone() if isinstance(x, torch.Tensor) else np.copy(x)
    y[:, 0] = x[:, 0] - x[:, 2] / 2  # top left x
    y[:, 1] = x[:, 1] - x[:, 3] / 2  # top left y
    y[:, 2] = x[:, 0] + x[:, 2] / 2  # bottom right x
    y[:, 3] = x[:, 1] + x[:, 3] / 2  # bottom right y
    return y


def xyxy2xywh(x):  # utils.general
    # Convert nx4 boxes from [x1, y1, x2, y2] to [x, y, w, h] where xy1=top-left, xy2=bottom-right
    y = x.clone() if isinstance(x, torch.Tensor) else np.copy(x)
    y[:, 0] = (x[:, 0] + x[:, 2]) / 2  # x center
    y[:, 1] = (x[:, 1] + x[:, 3]) / 2  # y center
    y[:, 2] = x[:, 2] - x[:, 0]  # width
    y[:, 3] = x[:, 3] - x[:, 1]  # height
    return y


def box_iou(box1, box2):  # utils.metrics
    # https://github.com/pytorch/vision/blob/master/torchvision/ops/boxes.py
    """
    Return intersection-over-union (Jaccard index) of boxes.
    Both sets of boxes are expected to be in (x1, y1, x2, y2) format.
    Arguments:
        box1 (Tensor[N, 4])
        box2 (Tensor[M, 4])
    Returns:
        iou (Tensor[N, M]): the NxM matrix containing the pairwise
            IoU values for every element in boxes1 and boxes2
    """

    def box_area(box):
        # box = 4xn
        return (box[2] - box[0]) * (box[3] - box[1])

    area1 = box_area(box1.T)
    area2 = box_area(box2.T)

    # inter(N,M) = (rb(N,M,2) - lt(N,M,2)).clamp(0).prod(2)
    inter = (torch.min(box1[:, None, 2:], box2[:, 2:]) - torch.max(box1[:, None, :2], box2[:, :2])).clamp(0).prod(2)
    return inter / (area1[:, None] + area2 - inter)  # iou = inter / (area1 + area2 - inter)


def non_max_suppression(prediction, conf_thres=0.25, iou_thres=0.45, classes=None, agnostic=False, multi_label=False,
                        labels=(), max_det=300, unknown=True):  # utils.general
    """Runs Non-Maximum Suppression (NMS) on inference results

    Returns:
         list of detections, on (n,6) tensor per image [xyxy, conf, cls]
    """
    nc = prediction.shape[2] - 5  # number of classes
    if unknown:
        xc = prediction[..., 4] > 0.2
    else:
        xc = prediction[..., 4] > conf_thres  # candidates

    # Settings
    min_wh, max_wh = 2, 4096  # (pixels) minimum and maximum box width and height
    max_nms = 30000  # maximum number of boxes into torchvision.ops.nms()
    time_limit = 10.0  # seconds to quit after
    redundant = True  # require redundant detections
    multi_label &= nc > 1  # multiple labels per box (adds 0.5ms/img)
    merge = False  # use merge-NMS

    t = time.time()
    output = [torch.zeros((0, 6), device=prediction.device)] * prediction.shape[0]
    for xi, x in enumerate(prediction):  # image index, image inference
        # Apply constraints
        # x[((x[..., 2:4] < min_wh) | (x[..., 2:4] > max_wh)).any(1), 4] = 0  # width-height
        x = x[xc[xi]]  # confidence

        # Cat apriori labels if autolabelling
        if labels and len(labels[xi]):
            l = labels[xi]
            v = torch.zeros((len(l), nc + 5), device=x.device)
            v[:, :4] = l[:, 1:5]  # box
            v[:, 4] = 1.0  # conf
            v[range(len(l)), l[:, 0].long() + 5] = 1.0  # cls
            x = torch.cat((x, v), 0)

        # If none remain process next image
        if not x.shape[0]:
            continue

        # Compute conf
        x[:, 5:] *= x[:, 4:5]  # conf = obj_conf * cls_conf

        # Box (center x, center y, width, height) to (x1, y1, x2, y2)
        box = xywh2xyxy(x[:, :4])

        # Detections matrix nx6 (xyxy, conf, cls)
        if multi_label:
            i, j = (x[:, 5:] > conf_thres).nonzero(as_tuple=False).T
            x = torch.cat((box[i], x[i, j + 5, None], j[:, None].float()), 1)
        else:  # best class only
            conf, j = x[:, 5:].max(1, keepdim=True)
            x = torch.cat((box, conf, j.float()), 1)[(conf.view(-1) > 0.2) if unknown else (conf.view(-1) > conf_thres)]

        # Filter by class
        if classes is not None:
            x = x[(x[:, 5:6] == torch.tensor(classes, device=x.device)).any(1)]

        # Apply finite constraint
        # if not torch.isfinite(x).all():
        #     x = x[torch.isfinite(x).all(1)]

        # Check shape
        n = x.shape[0]  # number of boxes
        if not n:  # no boxes
            continue
        elif n > max_nms:  # excess boxes
            x = x[x[:, 4].argsort(descending=True)[:max_nms]]  # sort by confidence

        # Batched NMS
        c = x[:, 5:6] * (0 if agnostic else max_wh)  # classes
        boxes, scores = x[:, :4] + c, x[:, 4]  # boxes (offset by class), scores
        i = torchvision.ops.nms(boxes, scores, iou_thres)  # NMS
        if i.shape[0] > max_det:  # limit detections
            i = i[:max_det]
        if merge and (1 < n < 3E3):  # Merge NMS (boxes merged using weighted mean)
            # update boxes as boxes(i,4) = weights(i,n) * boxes(n,4)
            iou = box_iou(boxes[i], boxes) > iou_thres  # iou matrix
            weights = iou * scores[None]  # box weights
            x[i, :4] = torch.mm(weights, x[:, :4]).float() / weights.sum(1, keepdim=True)  # merged boxes
            if redundant:
                i = i[iou.sum(1) > 1]  # require redundancy

        output[xi] = x[i]
        if unknown:
            output[0][..., 5] *= output[0][..., 4] >= conf_thres  # 신뢰도가 conf값보다 크면 그대로, 아니면 0으로 처리
            output[0][..., 5] += nc * (output[0][..., 4] < conf_thres)  # 신뢰도가 conf 값보다 작으면 클래스값을 곱해서 새로운 클래스 생성
        if (time.time() - t) > time_limit:
            print(f'WARNING: NMS time limit {time_limit}s exceeded')
            break  # time limit exceeded

    return output


def clip_coords(boxes, shape):
    # Clip bounding xyxy bounding boxes to image shape (height, width)
    if isinstance(boxes, torch.Tensor):  # faster individually
        boxes[:, 0].clamp_(0, shape[1])  # x1
        boxes[:, 1].clamp_(0, shape[0])  # y1
        boxes[:, 2].clamp_(0, shape[1])  # x2
        boxes[:, 3].clamp_(0, shape[0])  # y2
    else:  # np.array (faster grouped)
        boxes[:, [0, 2]] = boxes[:, [0, 2]].clip(0, shape[1])  # x1, x2
        boxes[:, [1, 3]] = boxes[:, [1, 3]].clip(0, shape[0])  # y1, y2


def scale_coords(img1_shape, coords, img0_shape, ratio_pad=None):
    # Rescale coords (xyxy) from img1_shape to img0_shape
    if ratio_pad is None:  # calculate from img0_shape
        gain = min(img1_shape[0] / img0_shape[0], img1_shape[1] / img0_shape[1])  # gain  = old / new
        pad = (img1_shape[1] - img0_shape[1] * gain) / 2, (img1_shape[0] - img0_shape[0] * gain) / 2  # wh padding
    else:
        gain = ratio_pad[0][0]
        pad = ratio_pad[1]

    coords[:, [0, 2]] -= pad[0]  # x padding
    coords[:, [1, 3]] -= pad[1]  # y padding
    coords[:, :4] /= gain
    clip_coords(coords, img0_shape)
    return coords


class Annotator:  # utils.plots
    if RANK in (-1, 0):
        check_font()  # download TTF if necessary

    # YOLOv5 Annotator for train/val mosaics and jpgs and detect/hub inference annotations
    def __init__(self, im, line_width=None, font_size=None, font='Arial.ttf', pil=False, example='abc'):
        assert im.data.contiguous, 'Image not contiguous. Apply np.ascontiguousarray(im) to Annotator() input images.'
        self.pil = pil or not is_ascii(example)
        if self.pil:  # use PIL
            self.im = im if isinstance(im, Image.Image) else Image.fromarray(im)
            self.draw = ImageDraw.Draw(self.im)
            self.font = check_font(font=font, size=font_size or max(round(sum(self.im.size) / 2 * 0.035), 12))
        else:  # use cv2
            self.im = im
        self.lw = line_width or max(round(sum(im.shape) / 2 * 0.003), 2)  # line width

    def box_label(self, box, label='', color=(128, 128, 128), txt_color=(255, 255, 255)):
        # Add one xyxy box to image with label
        if self.pil or not is_ascii(label):
            self.draw.rectangle(box, width=self.lw, outline=color)  # box
            if label:
                w, h = self.font.getsize(label)  # text width, height
                outside = box[1] - h >= 0  # label fits outside box
                self.draw.rectangle([box[0],
                                     box[1] - h if outside else box[1],
                                     box[0] + w + 1,
                                     box[1] + 1 if outside else box[1] + h + 1], fill=color)
                # self.draw.text((box[0], box[1]), label, fill=txt_color, font=self.font, anchor='ls')  # for PIL>8.0
                self.draw.text((box[0], box[1] - h if outside else box[1]), label, fill=txt_color, font=self.font)
        else:  # cv2
            p1, p2 = (int(box[0]), int(box[1])), (int(box[2]), int(box[3]))
            cv2.rectangle(self.im, p1, p2, color, thickness=self.lw, lineType=cv2.LINE_AA)
            if label:
                tf = max(self.lw - 1, 1)  # font thickness
                w, h = cv2.getTextSize(label, 0, fontScale=self.lw / 3, thickness=tf)[0]  # text width, height
                outside = p1[1] - h - 3 >= 0  # label fits outside box
                p2 = p1[0] + w, p1[1] - h - 3 if outside else p1[1] + h + 3
                cv2.rectangle(self.im, p1, p2, color, -1, cv2.LINE_AA)  # filled
                cv2.putText(self.im, label, (p1[0], p1[1] - 2 if outside else p1[1] + h + 2), 0, self.lw / 3, txt_color,
                            thickness=tf, lineType=cv2.LINE_AA)

    def rectangle(self, xy, fill=None, outline=None, width=1):
        # Add rectangle to image (PIL-only)
        self.draw.rectangle(xy, fill, outline, width)

    def text(self, xy, text, txt_color=(255, 255, 255)):
        # Add text to image (PIL-only)
        w, h = self.font.getsize(text)  # text width, height
        self.draw.text((xy[0], xy[1] - h + 1), text, fill=txt_color, font=self.font)

    def result(self):
        # Return annotated image as array
        return np.asarray(self.im)


class Colors:  # utils.plots
    # Ultralytics color palette https://ultralytics.com/
    def __init__(self):
        # hex = matplotlib.colors.TABLEAU_COLORS.values()
        hex = ('FF3838', 'FF9D97', 'FF701F', 'FFB21D', 'CFD231', '48F90A', '92CC17', '3DDB86', '1A9334', '00D4BB',
               '2C99A8', '00C2FF', '344593', '6473FF', '0018EC', '8438FF', '520085', 'CB38FF', 'FF95C8', 'FF37C7')
        self.palette = [self.hex2rgb('#' + c) for c in hex]
        self.n = len(self.palette)

    def __call__(self, i, bgr=False):
        c = self.palette[int(i) % self.n]
        return (c[2], c[1], c[0]) if bgr else c

    @staticmethod
    def hex2rgb(h):  # rgb order (PIL)
        return tuple(int(h[1 + i:1 + i + 2], 16) for i in (0, 2, 4))


colors = Colors()  # create instance


class DetThread(QThread):  # 쓰레드 정의
    send_img = pyqtSignal(np.ndarray)  # boxed image  처리 이미지
    send_raw = pyqtSignal(np.ndarray)  # magnified image  원본 이미지
    # send_statistic = pyqtSignal(dict)  # detecting result  처리 결과물
    send_frame = pyqtSignal(int)  # current lbl_frame  현재 프레임
    send_frames = pyqtSignal(int)  # total lbl_frame  총 프레임
    send_time = pyqtSignal(datetime)  # mtime  파일생성시간
    send_night = pyqtSignal(bool)  # noon/night  낮/밤 여부
    send_cnt = pyqtSignal(str)  # permanent statusbar to count/nf  우측하단 상태바 (현재파일번호/전체파일개수)
    send_det = pyqtSignal(int)  # number of detections  개체 수
    send_path = pyqtSignal(str)  # file path  파일 경로
    send_status = pyqtSignal(
        int)  # status (1:video is running, 2:video is stopped, 3:image  상태 (1:비디오 실행 중, 2:비디오 멈춤, 3:이미지)
    send_autocheck = pyqtSignal(int)  # judge category  카테고리 판별
    send_conf = pyqtSignal(tuple)  # max-conf, min-conf  최대 신뢰도, 최소 신뢰도
    send_disable = pyqtSignal(bool)  # all btn disable  모든 버튼 비활성
    send_finish = pyqtSignal(bool)  # when auto finish  자동분류 끝낱 을 때
    send_detect_path = pyqtSignal(str)  # 객체 인식 폴더 경로
    send_background_path = pyqtSignal(str)  # 배경 폴더 경로
    send_info = pyqtSignal(str)  # status info  현재 상태정보

    def __init__(self):
        super(DetThread, self).__init__()
        self.weights = './model/boargorani.pt'  # model path  모델 경로
        # self.weights = 0  # model
        self.source = './sample'  # dir  폴더 경로
        self.imgsz = [480, 480]  # inference size(pixel)  분할 개수
        self.conf_thres = 0.65  # 신뢰 임계값
        self.iou_thres = 0.45  # 교차범위 임계값
        self.save_txt = False  # 결과물 저장여부
        # self.nosave = True  # 저장여부
        self.classes = None  # 카테고리
        self.agnostic_nms = False  # 객체만 인식
        self.project = './runs/detect'  # 프로젝트 경로
        self.name = 'exp'  # 결과물 저장경로
        self.exist_ok = False  # 결과물 덮어쓰기
        self.line_thickness = 2  # 박스 굵기
        self.hide_labels = False  # 라벨 숨기기
        self.hide_conf = True  # 신뢰도 숨기기
        self.unknown = False  # under conf-thres to label unknown  신뢰임계값 미만 unknown 처리
        self.speed = 1  # lbl_frame delay(ms)  속도 지연값(밀리초)
        self.speed_up = False  # 2배속
        self.cond = QWaitCondition()  # thread control  쓰레드 컨트롤
        self.status = True  # thread pause  쓰레드 일시정지
        self.mutex = QMutex()  # thread lock  쓰레드 잠금
        self.auto_li = []  # category list  카테고리 목록
        self.conf_li = []  # confidence list  신뢰도 목록
        self.anicount = 0  # number of detection  개체 수
        self.auto = False  # 자동분류 여부
        self.error = './icon/error.JPG'
        self.device = ''

    def play(self):  # 재생 함수
        cv2.destroyAllWindows()  # 확대된 사진이 있을 경우 닫음
        self.status = True  # resume while loop  반복문 활성화
        self.send_status.emit(1)  # video is running
        self.cond.wakeAll()  # resume thread  쓰레드 활성화
        logger.info('play')

    def pause(self):  # 일시정지 함수
        cv2.destroyAllWindows()  # 확대된 사진이 있을 경우 닫음
        self.status = False  # pause while loop  반복문 비활성화
        self.send_status.emit(2)  # video is stopped
        if len(self.auto_li):  # If number of detection > 0  감지된 객체가 있다면
            self.send_autocheck.emit(
                Counter(self.auto_li).most_common()[0][0])  # send signal for most common category  가장 많이 감지된 카테고리
            self.send_conf.emit((max(self.conf_li), min(self.conf_li)))  # send signal for max, min conf  최대 신뢰도와 최소 신뢰도
        else:
            self.send_autocheck.emit(99)  # send default  미동정 처리
            self.send_conf.emit((0, 0))  # no detect  신뢰도 0, 0 처리
        logger.info('pause')

    def next(self):  # 다음 파일 불러오는 함수
        cv2.destroyAllWindows()  # 확대된 사진이 있을 경우 닫음
        if self.count < self.nf - 1:  # if current file num < total file num  만약 현재 파일번호 < 전체 파일 개수면
            self.auto_li.clear()  # reset animal list  동물목록 초기화
            self.conf_li.clear()  # reset confidence list  신뢰도목록 초기화
            self.anicount = 0  # reset number of detection 개체 수 초기화
            self.count += 1  # 다음 파일번호
            self.path = self.files[self.count]  # 파일 경로
            self.send_path.emit(self.path)
            self.send_cnt.emit(f'{self.count + 1}/{self.nf}')  # 파일번호 / 전체 파일 개수
            if self.video_flag[self.count]:  # if video:  만약 영상이면  [False, False, False, True, True, True]
                self.frame = 0  # reset lbl_frame  프레임 초기화
                self.mode = 'video'  # 영상 처리
                self.vid_cap = cv2.VideoCapture(self.path)  # 영상 읽기
                self.frames = int(self.vid_cap.get(cv2.CAP_PROP_FRAME_COUNT))  # total lbl_frame  총 프레임
            self.send_night.emit(False)  # reset noon/night(to noon)  주야간 초기화 (주간처리)
            self.status = True  # resume while loop  반복문 활성화
            self.send_det.emit(0)  # reset number of detection  개체 수 초기화
            if self.mode == 'video':  # 비디오면
                self.send_status.emit(1)  # video is running
            self.cond.wakeAll()  # resume thread  쓰레드 활성화
            logger.info('next')
        else:  # last file  마지막 파일이면
            QTest.qWait(100)  # delay 0.1 sec  0.1초 지연

    def prev(self):  # 이전 파일 불러오는 함수
        cv2.destroyAllWindows()  # 확대된 사진이 있을 경우 닫음
        self.send_det.emit(0)  # reset number of detection  개체 수 초기화
        self.frame = 0  # reset lbl_frame  프레임 초기화
        self.auto_li.clear()  # reset animal list  동물목록 초기화
        self.conf_li.clear()  # reset confidence list  신뢰도목록 초기화
        self.anicount = 0  # reset number of detection 개체 수 초기화
        if self.count > 0:  # prevent minus count error  첫 파일일 경우 처리
            try:
                self.vid_cap.release()
            except Exception:
                pass
            finally:
                self.count -= 1  # 이전 파일번호
                self.send_cnt.emit(f'{self.count + 1}/{self.nf}')  # 파일번호 / 전체 파일 개수
        self.path = self.files[self.count]  # 파일 경로
        self.send_path.emit(self.path)
        if self.video_flag[self.count]:  # video
            self.vid_cap = cv2.VideoCapture(self.path)  # 영상 읽기
            self.frames = int(self.vid_cap.get(cv2.CAP_PROP_FRAME_COUNT))  # total lbl_frame  총 프레임
            # self.send_frames.emit(self.frames)  # addhee
            self.send_night.emit(False)  # reset noon/night(to noon)  주야간 초기화 (주간처리)
            self.status = True  # resume while loop  반복문 활성화
            self.send_status.emit(1)  # video is running
            self.cond.wakeAll()  # resume thread  쓰레드 활성화
            self.status = False  # pause while loop(to run once)  반복문 비활성화(첫 프레임만 처리)
            self.send_status.emit(2)  # video is stopped
        else:  # image
            self.mode = 'image'
            self.send_night.emit(False)  # reset noon/night(to noon)  주야간 초기화 (주간처리)
            self.play()  # run thread once  쓰레드 1번 돌림
            self.send_frames.emit(0)  # total lbl_frame to 0  총 프레임 0으로 처리
        logger.info('prev')

    def move(self, num):  # 파일 이동
        cv2.destroyAllWindows()
        try:
            self.vid_cap.release()
        except Exception:
            pass
        self.send_det.emit(0)  # reset number of detection  개체 수 초기화
        self.auto_li.clear()  # reset animal list  동물목록 초기화
        self.conf_li.clear()  # reset confidence list  신뢰도목록 초기화
        self.anicount = 0  # reset number of detection 개체 수 초기화
        self.count = num  # 원하는 파일 번호로 세팅
        if self.video_flag[num]:  # if video:  만약 영상이면  [False, False, False, False, False, False, True, True, True]
            self.frame = 0  # reset lbl_frame  프레임 초기화
            self.mode = 'video'
            self.path = self.files[self.count]  # 파일 경로
            self.vid_cap = cv2.VideoCapture(self.path)  # 영상 읽기
            self.frames = int(self.vid_cap.get(cv2.CAP_PROP_FRAME_COUNT))  # total lbl_frame  총 프레임
        else:  # image
            self.mode = 'image'
            self.path = self.files[self.count]  # 파일 경로
            self.send_frames.emit(0)  # total lbl_frame to 0  총 프레임 0으로 처리
        self.send_path.emit(self.path)
        self.send_cnt.emit(f'{self.count + 1}/{self.nf}')  # 파일번호 / 전체 파일 개수
        self.send_night.emit(False)  # reset noon/night(to noon)  주야간 초기화 (주간처리)
        self.status = True  # resume while loop  반복문 활성화
        if self.mode == 'video':
            self.send_status.emit(1)  # video is running
        self.cond.wakeAll()  # resume thread  쓰레드 활성화
        logger.info('move')

    def speed_slow(self, sp_text):  # 속도 조절 함수
        self.speed_up = False
        if sp_text == 'Very Slow':
            self.speed = 200  # 프레임당 0.2초 지연
        elif sp_text == 'Slow':
            self.speed = 100  # 프레임당 0.1초 지연
        elif sp_text == 'Normal':
            self.speed = 1  # 지연 없음
        elif sp_text == 'Fast':
            self.speed_up = True  # 2배속
            self.speed = 1

    def model(self, model):  # "멧돼지와 고라니", "멧돼지와 고라니와 산양", "사용 안함"
        if model == '사용 안함':
            self.weights = 0
        elif model == '멧돼지와 고라니':
            self.weights = './model/boargorani.pt'
        elif model == '멧돼지와 고라니와 산양':
            self.weights = './model/boargoranigoral.pt'
        logger.info(self.weights)

    def imgsz_opt(self, imgsz_text):  # inference size(pixel) 분할 개수
        if imgsz_text == '1280':
            self.imgsz = [1280, 1280]
        elif imgsz_text == '960':
            self.imgsz = [960, 960]
        elif imgsz_text == '640':
            self.imgsz = [640, 640]
        elif imgsz_text == '480':
            self.imgsz = [480, 480]
        elif imgsz_text == '320':
            self.imgsz = [320, 320]
        elif imgsz_text == '128':
            self.imgsz = [128, 128]

    def stopping(self):  # 쓰레드 멈춤 함수
        self.stop = True
        if not self.status:
            self.status = True
            self.cond.wakeAll()

    @torch.no_grad()  # detect.py
    def run(self):
        ff = np.fromfile('./icon/load.JPG', np.uint8)  # 모델을 불러오는 중
        img = cv2.imdecode(ff, cv2.IMREAD_COLOR)
        self.send_img.emit(img)
        self.send_disable.emit(True)
        # save_img = not self.nosave
        imgsz = self.imgsz

        # Directories
        save_dir = increment_path(Path(self.project) / self.name, exist_ok=self.exist_ok)  # increment run
        if self.save_txt:
            (save_dir / 'labels' if self.save_txt else save_dir).mkdir(parents=True, exist_ok=True)  # make dir

        # Initialize
        if self.weights:  # 모델을 선택하면
            device, info = select_device(self.device)  # gpu or cpu

            # Load model
            self.send_info.emit('딥러닝 모델을 불러오는 중...')
            model, device, info = attempt_load(self.weights, map_location=device, info=info)
            stride = int(model.stride.max())  # model stride  커널 보폭(32)
            self.names = model.module.names if hasattr(model, 'module') else model.names  # get class names
            self.names.append('Unknown')  # unknown 추가

            # Run inference
            if device.type != 'cpu':  # if gpu
                model(torch.zeros(1, 3, *imgsz).to(device).type_as(next(model.parameters())))  # run once
            self.send_info.emit(info)
        else:  # 모델이 없으면
            self.send_info.emit('영상인식 AI를 활용하고 있지 않습니다.')
        dt = [0.0, 0.0, 0.0]  # time check
        p = str(Path(self.source).resolve())  # os-agnostic absolute path
        files = natsorted(glob(os.path.join(p, '*.*')))  # dir
        images = [x for x in files if x.split('.')[-1].lower() in IMG_FORMATS]
        videos = [x for x in files if x.split('.')[-1].lower() in VID_FORMATS]
        ni, nv = len(images), len(videos)
        self.files = images + videos  # list of files
        self.nf = ni + nv  # number of files
        self.video_flag = [False] * ni + [True] * nv
        self.mode = 'image'
        if any(videos):  # 영상이 있다면
            self.frame = 0
            self.vid_cap = cv2.VideoCapture(videos[0])
            self.frames = int(self.vid_cap.get(cv2.CAP_PROP_FRAME_COUNT))
        else:
            self.vid_cap = None
        self.count = 0
        self.path = self.files[self.count]
        self.send_path.emit(self.path)  # path signal
        if not ni or self.auto:  # First file is video or automode
            self.send_status.emit(1)  # video is running
        self.send_night.emit(False)  # reset noon/night(to noon)  주야간 초기화 (주간처리)
        self.send_cnt.emit(f'1/{self.nf}')  # 1 / 총 파일 개수 표기
        self.stop = False
        if self.auto:
            ff = np.fromfile('./icon/auto.JPG', np.uint8)  # 자동모드 처리중 표시
            img = cv2.imdecode(ff, cv2.IMREAD_COLOR)
            self.send_img.emit(img)
        while True:  # 처리 시작
            self.mutex.lock()  # 반복문 잠금
            if not self.status:  # 일시정지 요청시
                self.cond.wait(self.mutex)  # 일시정지
            if self.stop:  # 쓰레드 종료
                self.mutex.unlock()
                self.terminate()
            if self.video_flag[self.count]:  # read video
                self.mode = 'video'
                ret_val, im0s = self.vid_cap.read()
                if not ret_val:  # 프레임을 못읽는 에러가 나면
                    if (self.frames == 0) and (self.frame < 1):
                        ff = np.fromfile(self.error, np.uint8)
                        im0s = cv2.imdecode(ff, cv2.IMREAD_COLOR)
                    elif self.frame < self.frames:  # 영상이 남아있으면 해당 프레임 넘김
                        self.frame += 1
                        self.vid_cap.set(cv2.CAP_PROP_POS_FRAMES, self.frame)  # set video to current lbl_frame
                        self.mutex.unlock()
                        continue
                    else:  # 마지막 프레임이면
                        if self.auto:  # auto에서 영상이 끝났을 때 다음영상으로
                            self.vid_cap.release()
                            self.count += 1
                            self.path = self.files[self.count]
                            self.send_path.emit(self.path)
                            self.frame = 0
                            self.vid_cap = cv2.VideoCapture(self.path)
                            self.frames = int(self.vid_cap.get(cv2.CAP_PROP_FRAME_COUNT))
                            ret_val, im0s = self.vid_cap.read()
                        else:  # 수동이면 멈추도록
                            self.status = False
                            self.send_status.emit(2)
                            self.mutex.unlock()
                            continue
                self.frame += 1
                if (self.frame % 2 == 0) and self.speed_up:  # 배속 상태이고 프레임이 짝수이면
                    # self.send_frames.emit(self.frames)  # addhee
                    self.send_frame.emit(self.frame)
                    if self.frame == self.frames:  # 마지막 프레임이면
                        self.status = False  # 멈춤
                        self.send_status.emit(2)  # video is stopped
                    self.mutex.unlock()
                    continue
                cv2.waitKey(self.speed)  # delay  지연시간
                s = f'video {self.count + 1}/{self.nf} ({self.frame}/{self.frames}) {self.path}: '
            else:  # read image
                try:
                    ff = np.fromfile(self.path, np.uint8)  # 한글 인식
                except FileNotFoundError:
                    ff = np.fromfile(self.error, np.uint8)
                else:
                    mtime = os.path.getmtime(self.path)  # file maketime  파일 생성시간
                    self.timedata = datetime.fromtimestamp(mtime)
                    self.send_time.emit(self.timedata)
                im0s = cv2.imdecode(ff, cv2.IMREAD_COLOR)  # BGR
                s = f'image {self.count + 1}/{self.nf} {self.path}: '
            imr = im0s.copy()
            h = imr.shape[0]
            if np.array_equal(imr[h // 5:h // 5 * 4, :, 0], imr[h // 5:h // 5 * 4, :, 1]):  # judge noon/night  주야간 판단
                self.send_night.emit(True)  # 야간처리
            # statistic_dic = {name: 0 for name in self.names}
            t1 = time_sync()
            if self.weights:  # 모델사용
                img = letterbox(im0s, self.imgsz, stride=stride, auto=True)[0]  # padded resize
                img = img.transpose((2, 0, 1))[::-1]  # HWC to CHW, BGR to RGB
                img = np.ascontiguousarray(img)
                img = np.array(img)
                img = torch.from_numpy(img).to(device)
                img = img.float()
                img /= 255.0  # 0 - 255 to 0.0 - 1.0
                if len(img.shape) == 3:
                    img = img[None]  # expand for batch dim
                t2 = time_sync()
                dt[0] += t2 - t1

                # Inference
                visualize = False
                augment = False
                pred = model(img, augment=augment, visualize=visualize)[0]

                # NMS
                pred = non_max_suppression(pred, self.conf_thres, self.iou_thres, self.classes, self.agnostic_nms,
                                           max_det=100, unknown=self.unknown)
                t3 = time_sync()
                dt[1] += t3 - t2
                dt[2] += time_sync() - t3
                if self.auto:  # auto  자동분류모드이면
                    if self.mode == 'video':  # video  영상이면
                        if not len(pred[0]):  # no detect  미감지면
                            if self.frame == self.frames:  # last lbl_frame  마지막 프레임이면
                                self.send_background_path.emit(self.path)  # to background  배경처리
                                if self.count == self.nf - 1:  # last file  마지막 파일이면
                                    self.status = False  # stop
                                    self.send_status.emit(2)  # video is stopped
                                    self.mutex.unlock()
                                    self.send_finish.emit(True)  # 자동분류 종료
                                else:  # if not last file  마지막 파일이 아니면
                                    self.mutex.unlock()
                                    self.next()  # 다음파일
                            else:  # if not last lbl_frame  마지막 프레임이 아니면
                                self.mutex.unlock()  # next lbl_frame  다음 프레임
                        else:  # if detect  객체가 감지되면
                            self.send_detect_path.emit(self.path)  # to detect  감지처리
                            if self.count == self.nf - 1:  # last file  마지막 파일이면
                                self.status = False  # stop  멈춤
                                self.send_status.emit(2)  # video is stopped
                                self.mutex.unlock()
                                self.send_finish.emit(True)  # 자동분류 종료
                            else:  # if not last file  # 마지막 파일이 아니면
                                self.mutex.unlock()
                                self.next()  # 다음파일
                    else:  # image
                        if len(pred[0]):  # if detect  감지되면
                            self.send_detect_path.emit(self.path)  # to detect  감지처리
                        else:  # no detect  미감지면
                            self.send_background_path.emit(self.path)  # to background  배경처리
                        self.count += 1
                        self.send_cnt.emit(f'{self.count + 1}/{self.nf}')
                        if self.count == self.nf:  # last file  마지막 파일이면
                            self.status = False  # stop  멈춤
                            self.send_status.emit(2)  # video is stopped
                            self.mutex.unlock()
                            self.send_finish.emit(True)  # 자동분류 종료
                        else:
                            self.path = self.files[self.count]
                            self.mutex.unlock()
                else:  # manual  수동검수모드
                    # Process predictions
                    for det in pred:  # per image  이미지마다
                        p, im0 = self.path, im0s.copy()
                        p = Path(p)  # to Path
                        txt_path = str(save_dir / 'labels' / p.stem) + (
                            '' if self.mode == 'image' else f'_{self.frame}')  # img.txt  딥러닝 결과 txt 현재 막아둠
                        s += '%gx%g ' % img.shape[2:]  # print string
                        gn = torch.tensor(im0.shape)[[1, 0, 1, 0]]  # normalization gain whwh
                        annotator = Annotator(im0, line_width=self.line_thickness, example=str(self.names))
                        if len(det):  # 객체가 있으면
                            if self.anicount < len(det):  # 이전까지 감지된 최대 개체 수 보다 현재 개체 수가 더 크면
                                self.anicount = len(det)  # 개체 수 갱신
                                self.send_det.emit(self.anicount)  # animal count signal  개체 수 시그널
                            # Rescale boxes from img_size to im0 size
                            det[:, :4] = scale_coords(img.shape[2:], det[:, :4], im0.shape).round()  # 0,0,1,1 을 원래 사이즈로

                            # Print results
                            for c in det[:, -1].unique():
                                n = (det[:, -1] == c).sum()  # detections per class
                                s += f"{n} {self.names[int(c)]}{'s' * (n > 1)}, "  # add to string

                            # Write results
                            for *xyxy, conf, cls in det:
                                conf = f'{conf:.2f}'
                                if self.save_txt:  # Write to file
                                    xywh = (xyxy2xywh(torch.tensor(xyxy).view(1, 4)) / gn).view(
                                        -1).tolist()  # normalized xywh
                                    line = (cls, *xywh)  # label format
                                    with open(txt_path + '.txt', 'a') as f:
                                        f.write(('%g ' * len(line)).rstrip() % line + '\n')

                                # Add bbox to image
                                c = int(cls)  # integer class
                                # statistic_dic[self.names[c]] += 1
                                self.auto_li.append(c)  # 동물목록 추가
                                self.conf_li.append(conf)  # 신뢰값 목록 추가
                                label = None if self.hide_labels else (
                                    self.names[c] if self.hide_conf else f'{self.names[c]} {conf}')
                                annotator.box_label(xyxy, label, color=colors(c, True))  # 사진에 라벨 작성

                        # Print time (inference-only)
                        print(f'{s}Done. ({t3 - t2:.3f}s)')  # 1221
                        # cv2.waitKey(1)  # 1 millisecond
                        # Stream results
                        im0 = annotator.result()  # 처리 사진

                        # Save results (image with detections)
                        # if save_img:
                        #     if self.mode == 'image':
                        #         cv2.imwrite(save_path, im0)
                        #     else:  # 'video' or 'stream'
                        #         if vid_path[i] != save_path:  # new video
                        #             vid_path[i] = save_path
                        #             if isinstance(vid_writer[i], cv2.VideoWriter):
                        #                 vid_writer[i].release()  # release previous video writer
                        #             if self.vid_cap:  # video
                        #                 fps = self.vid_cap.get(cv2.CAP_PROP_FPS)
                        #                 w = int(self.vid_cap.get(cv2.CAP_PROP_FRAME_WIDTH))
                        #                 h = int(self.vid_cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
                        #             else:  # stream
                        #                 fps, w, h = 30, im0.shape[1], im0.shape[0]
                        #                 save_path += '.mp4'
                        #             vid_writer[i] = cv2.VideoWriter(save_path, cv2.VideoWriter_fourcc(*'mp4v'), fps, (w, h))
                        #         vid_writer[i].write(im0)

                        self.send_raw.emit(imr)  # 원본사진
                        self.send_img.emit(im0)  # 처리사진
                        # self.send_statistic.emit(statistic_dic)  # detecting 결과

                        if self.mode == 'video':
                            self.send_frames.emit(self.frames)
                            self.send_frame.emit(self.frame)
                            if self.frame == self.frames:  # last lbl_frame  마지막 프레임이면
                                if len(self.auto_li):  # If number of detection > 0  감지된 객체가 있다면
                                    self.send_autocheck.emit(Counter(self.auto_li).most_common()[0][
                                                                 0])  # send signal for most common category  가장 많이 감지된 카테고리
                                    self.send_conf.emit((max(self.conf_li),
                                                         min(self.conf_li)))  # send signal for max, min conf  최대 신뢰도와 최소 신뢰도
                                else:
                                    self.send_autocheck.emit(99)  # send default  미동정 처리
                                    self.send_conf.emit((0, 0))  # no detect  신뢰도 0, 0 처리
                                self.status = False  # pause while loop  반복문 비활성화
                                self.send_status.emit(2)  # video is stopped
                        else:  # image
                            if len(self.auto_li):  # If number of detection > 0  감지된 객체가 있다면
                                self.send_autocheck.emit(Counter(self.auto_li).most_common()[0][
                                                             0])  # send signal for most common category  가장 많이 감지된 카테고리
                                self.send_conf.emit((max(self.conf_li),
                                                     min(self.conf_li)))  # send signal for max, min conf  최대 신뢰도와 최소 신뢰도
                            else:
                                self.send_autocheck.emit(99)  # send default  미동정 처리
                                self.send_conf.emit((0, 0))  # no detect  신뢰도 0, 0 처리
                            self.send_status.emit(3)  # image
                            self.status = False  # pause while loop  반복문 비활성화
                    self.mutex.unlock()  # 반복문 해제

            else:  # 딥러닝 사용x
                self.send_raw.emit(imr)  # 원본사진
                self.send_img.emit(imr)  # 원본사진
                if self.mode == 'video':
                    self.send_frames.emit(self.frames)  # total lbl_frame  총 프레임
                    self.send_frame.emit(self.frame)  # current lbl_frame  현재 프레임
                    if self.frame == self.frames:  # last lbl_frame  마지막 프레임
                        self.status = False  # pause while loop  반복문 비활성화
                        self.send_status.emit(2)  # video is stopped
                else:  # image
                    self.send_status.emit(3)  # image
                    self.status = False  # pause while loop  반복문 비활성화
                self.mutex.unlock()  # 반복문 해제


class MyApp(QMainWindow):  # 메인윈도우 정의
    def __init__(self):
        super().__init__()
        # self.setWindowFlags(Qt.WindowCloseButtonHint)  # 플래그
        self.cent_widget = CentWidget()  # 위젯 클래스 상속
        self.setCentralWidget(self.cent_widget.tabs)  # 위젯 배치
        self.setFocusPolicy(Qt.StrongFocus)  # 포커싱(키보드 컨트롤을 위해)
        statusbar = self.statusBar()
        # statusbar.addPermanentWidget(QProgressBar())
        statusbar.addPermanentWidget(self.cent_widget.lbl_info)  # gpu info
        statusbar.addPermanentWidget(QLabel('   Copyrightⓒ2021 BIGLeader All rights reserved     '))
        statusbar.addPermanentWidget(self.cent_widget.lbl_cnt)  # file num 파일 번호
        self.setGeometry(20, 50, int(0.8 * width), int(0.8 * height))  # xywh
        self.setWindowTitle('Oasis')
        self.setWindowIcon(icon)
        menubar = self.menuBar()
        filemenu = QMenu("메뉴", self)
        makermenu = QMenu('제작자', self)
        menubar.addMenu(filemenu)
        # menubar.addMenu(makermenu)
        self.sourcemenu = QAction('실행', self)
        self.sourcemenu.setShortcut('F5')
        self.sourcemenu.triggered.connect(self.cent_widget.source)
        self.stopmenu = QAction('실행 중지', self)
        self.stopmenu.triggered.connect(self.stop)
        helpmenu = QAction("도움말", self)
        helpmenu.triggered.connect(self.helpbar)
        filemenu.addAction(self.sourcemenu)
        filemenu.addAction(self.stopmenu)
        filemenu.addAction(helpmenu)
        filemenu.aboutToShow.connect(self.signal)
        creditmenu = QAction('제작자', self)
        # questionmenu = QAction('문의하기', self)
        makermenu.addAction(creditmenu)
        creditmenu.triggered.connect(self.credit)
        current_version = QLabel('Oasis v' + version + '   ')
        menubar.setCornerWidget(current_version)
        self.cent_widget.det_thread.send_info.connect(self.info)
        self.show()
        self.showMaximized()
        if not self.cent_widget.save_status:  # 캐시 파일이 존재하면
            QMessageBox.warning(self, '복원됨', '비정상 종료로 인한 데이터가 복구되었습니다.', QMessageBox.Yes)
        self.cent_widget.source()  # 폴더 선택

    def credit(self):  # 제작자
        self.c = PaintPicture(3)  # 꺼지지 않게 변수 할당

    def helpbar(self):  # 도움말
        self.helper = PaintPicture(self.cent_widget.tabs.currentIndex())

    def info(self, info):  # gpu info
        self.cent_widget.lbl_info.setText(info)

    def signal(self):  # 쓰레드 상태 판단
        if self.cent_widget.det_thread.isRunning():
            self.sourcemenu.setEnabled(False)
            self.stopmenu.setEnabled(True)
        else:
            self.sourcemenu.setEnabled(True)
            self.stopmenu.setEnabled(False)

    def stop(self):  # 쓰레드 정지함수 호출
        self.cent_widget.stop()

    def keyPressEvent(self, e):  # 단축키 정의
        if e.key() == Qt.Key_F:  # press 'F' key
            self.showFullScreen()  # set app full screen
        elif e.key() == Qt.Key_M:  # press 'M' key
            self.showMaximized()  # set app size to maximize
        elif e.key() == Qt.Key_N:  # press 'N' key
            self.showNormal()  # set app size to normal
        elif e.key() == Qt.Key_P:  # press 'P' key
            if self.cent_widget.btn_pp.isEnabled():  # if btn play/pause enable
                self.cent_widget.btn_pp.click()  # btn play/pause click
        elif e.key() == Qt.Key_A:  # press 'A' key
            if self.cent_widget.btn_prev.isEnabled():  # if btn prev enable
                self.cent_widget.btn_prev.click()  # btn prev click
        elif e.key() == Qt.Key_D:  # press 'D' key
            if self.cent_widget.btn_next.isEnabled():  # if btn next enable
                self.cent_widget.btn_next.click()  # btn next click
        elif e.key() == Qt.Key_S:  # press 'S' key
            if self.cent_widget.btn_submit.isEnabled():  # if btn submit enable
                self.cent_widget.btn_submit.click()  # btn submit click
        elif e.key() == Qt.Key_X:  # press 'X' key
            if self.cent_widget.btn_rm_file.isEnabled():
                self.cent_widget.btn_rm_file.click()  # btn 파일삭제 click
        elif e.key() == Qt.Key_Escape:
            self.setFocus()

    def closeEvent(self, event):  # 프로그램 종료 정의
        if not self.cent_widget.save_status:  # 상태가 저장되지 않으면
            reply2 = QMessageBox.warning(self, '저장되지 않은 데이터',  # 경고 출력
                                         'Result 탭의 결과물이 저장되지 않았습니다.\n저장되지 않은 데이터는 사라집니다.\n그래도 종료하시겠습니까?',
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply2 == QMessageBox.No:
                event.ignore()
        if self.cent_widget.save_status or reply2 == QMessageBox.Yes:  # 저장된 상태이거나 이전질문에서 확인을 누르면
            reply = QMessageBox.question(self, '나가기', '정말로 종료하시겠습니까?',
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                try:
                    os.remove(CONFIG_DIR / 'cache.csv')
                except FileNotFoundError:
                    pass
                event.accept()
            else:
                event.ignore()


class PaintPicture(QDialog):  # 새 창(도움말 등)
    def __init__(self, signal):
        super().__init__()
        self.signal = signal
        self.setWindowTitle('Help')
        self.setWindowIcon(icon)
        self.setGeometry(width // 10, height // 10, width // 10 * 8, height // 10 * 8)
        self.setWindowFlags(Qt.WindowCloseButtonHint)

        if self.signal == 0:  # main tab
            pixmap = QPixmap('./icon/help1.png')
        elif self.signal == 1:  # option tab
            pixmap = QPixmap('./icon/help2.png')
        elif self.signal == 2:  # result tab
            pixmap = QPixmap('./icon/help3.png')
        elif self.signal == 3:  # credit
            pixmap = QPixmap('./icon/credit.JPG')
            self.setWindowTitle('Credit')

        lbl_img = QLabel()
        lbl_img.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        lbl_img.setScaledContents(True)
        lbl_img.setPixmap(pixmap)

        vbox = QVBoxLayout()
        vbox.addWidget(lbl_img)
        self.setLayout(vbox)
        self.show()


class PhotoViewer(QGraphicsView):
    def __init__(self, parent):
        super(PhotoViewer, self).__init__(parent)
        self._zoom = 0
        self._empty = True
        self._scene = QGraphicsScene(self)
        self._photo = QGraphicsPixmapItem()
        self._scene.addItem(self._photo)
        self.setScene(self._scene)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setBackgroundBrush(QBrush(QColor(30, 30, 30)))
        self.setFrameShape(QFrame.NoFrame)

    def hasPhoto(self):
        return not self._empty

    def fitInView(self, scale=True):
        rect = QRectF(self._photo.pixmap().rect())
        if not rect.isNull():
            self.setSceneRect(rect)
            if self.hasPhoto():
                unity = self.transform().mapRect(QRectF(0, 0, 1, 1))
                self.scale(1 / unity.width(), 1 / unity.height())
                viewrect = self.viewport().rect()
                scenerect = self.transform().mapRect(rect)
                factor = min(viewrect.width() / scenerect.width(),
                             viewrect.height() / scenerect.height())
                self.scale(factor, factor)
            self._zoom = 0

    def setPhoto(self, pixmap=None):
        self._zoom = 0
        if pixmap and not pixmap.isNull():
            self._empty = False
            self.setDragMode(QGraphicsView.ScrollHandDrag)
            self._photo.setPixmap(pixmap)
        else:
            self._empty = True
            self.setDragMode(QGraphicsView.NoDrag)
            self._photo.setPixmap(QPixmap())
        self.fitInView()

    def wheelEvent(self, event):
        if self.hasPhoto():
            if event.angleDelta().y() > 0:
                factor = 1.25
                self._zoom += 1
            else:
                factor = 0.8
                self._zoom -= 1
            if self._zoom > 0:
                self.scale(factor, factor)
            elif self._zoom == 0:
                self.fitInView()
            else:
                self._zoom = 0

    def toggleDragMode(self):
        if self.dragMode() == QGraphicsView.ScrollHandDrag:
            self.setDragMode(QGraphicsView.NoDrag)
        elif not self._photo.pixmap().isNull():
            self.setDragMode(QGraphicsView.ScrollHandDrag)

    def mousePressEvent(self, event):
        if self._photo.isUnderMouse():
            self.photoClicked.emit(self.mapToScene(event.pos()).toPoint())
        super(PhotoViewer, self).mousePressEvent(event)


class ReadOnlyDelegate(QStyledItemDelegate):  # table disable  셀 비활성화
    def createEditor(self, parent, option, index):
        return


class CentWidget(QWidget):  # 위젯정의
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon('./icon/logo2.png'))
        # self.setFocusPolicy(Qt.NoFocus)

        self.det_thread = DetThread()  # 쓰레드 클래스 상속
        self.det_thread.send_raw.connect(self.mag)  # magnify  확대
        self.det_thread.send_img.connect(lambda x: self.show_image(x, self.img2))  # im0 (np.ndarray)
        self.det_thread.send_frame.connect(self.send_frame)  # current lbl_frame  현재 프레임
        self.det_thread.send_frames.connect(self.send_frames)  # 총 프레임 total lbl_frame
        self.det_thread.send_time.connect(self.filetime)  # file maketime 파일 생성시간
        self.det_thread.send_night.connect(self.night)  # noon/night  주/야간
        self.det_thread.send_cnt.connect(self.cnt)  # current file num/number of files  파일번호/총 파일개수
        self.det_thread.send_det.connect(self.animalcount)  # number of detections  개체 수
        self.det_thread.send_path.connect(self.path)  # 경로
        self.det_thread.send_status.connect(
            self.status)  # status (1:video is running, 2:video is stopped, 3:image  상태 (1:비디오 실행 중, 2:비디오 멈춤, 3:이미지)
        self.det_thread.send_autocheck.connect(self.autocheck)  # best category  최빈 카테고리
        self.det_thread.send_conf.connect(self.show_conf)  # max-conf, min-conf  최대, 최소 신뢰값
        self.det_thread.send_disable.connect(self.disable)  # 버튼 비활성화
        self.det_thread.send_finish.connect(self.auto_finish)  # 자동분류 종료
        self.det_thread.send_background_path.connect(self.background)  # 배경 경로
        self.det_thread.send_detect_path.connect(self.detect)  # 감지 경로

        self.app_data = check_appdata(file='app_data.yaml')  # table header, animal list
        self.animal_db = check_animal_db(file='animal_db.yaml')
        df = check_cache(file='cache.csv')
        if df is None:
            self.cache = pd.DataFrame(columns=self.app_data['columns'])
            self.save_status = True
        else:
            self.cache = df
            self.save_status = False
        del df

        self.background_list = []
        self.detect_list = []
        self.animal_db_dir = CONFIG_DIR / 'animal_db.yaml'
        self.app_data_dir = CONFIG_DIR / 'app_data.yaml'
        self.cache_dir = CONFIG_DIR / 'cache.csv'
        self.hours = 0

        gridmain = QGridLayout()
        gridopt = QGridLayout()
        gridtable = QGridLayout()

        font = QFont()
        font.setBold(True)
        font.setPointSize(10)

        big_font = QFont()
        big_font.setPointSize(15)
        big_font.setBold(True)

        # source : dir
        self.lbl_source = QLabel(os.getcwd(), self)
        # self.lbl_source = QLabel('D:/project/NationalPark_upgrade/PYQT5/yolov5_master/sample')  # 고정값
        self.lbl_source.setFont(font)
        self.lbl_source.setStyleSheet('background-color: #FFFFFF')
        self.lbl_source.setStatusTip("분류/감지할 '폴더'를 설정합니다. *경로상에 한글이 들어가면 오류가 발생할 수 있음!")
        self.lbl_source.setToolTip("분류/감지할 '폴더'를 설정합니다.\n*경로상에 한글이 들어가면 오류가 발생할 수 있음!")
        self.lbl_source.setFixedHeight(60)
        # self.lbl_source.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.btn_source = QPushButton('실행(폴더 선택)', self)
        self.btn_source.setFont(font)
        self.btn_source.setStatusTip("분류/감지할 '폴더'를 설정합니다. *경로상에 한글이 들어가면 오류가 발생할 수 있음!")
        self.btn_source.setToolTip("분류/감지할 '폴더'를 설정합니다.\n*경로상에 한글이 들어가면 오류가 발생할 수 있음!")
        self.btn_source.setStyleSheet("background-image: url(./icon/background.jpg);")
        self.btn_source.setFixedHeight(60)
        # self.btn_source.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.btn_source.clicked.connect(self.source)

        # imgsz : inference size (pixels)
        self.lbl_imgsz = QLabel('사진 분할크기', self)
        self.lbl_imgsz.setFont(font)
        self.lbl_imgsz.setStyleSheet('background-color: #FFFFFF')
        self.lbl_imgsz.setStatusTip('모델이 처리할 이미지 크기를 설정합니다. 값이 높을수록 이미지당 처리시간이 증가합니다.')
        self.lbl_imgsz.setToolTip('모델이 처리할 이미지 크기를 설정합니다.\n값이 높을수록 이미지당 처리시간이 증가합니다.')
        # btn_imgsz = QComboBox(self)
        # btn_imgsz.setFont(font)
        # [btn_imgsz.addItem(i) for i in ["1280", "960", "640", "480", "320", "128"]]
        # btn_imgsz.setCurrentIndex(3)
        # btn_imgsz.activated[str].connect(self.det_thread.imgsz_opt)
        # btn_imgsz = QPushButton('Image Size', self)
        # btn_imgsz.setFont(font)
        # btn_imgsz.setStatusTip('Set inference size (pixels) 128~1280')
        # btn_imgsz.setToolTip('Set inference size (pixels)\n128~1280')
        # btn_imgsz.clicked.connect(self.imgsz)

        # conf_thres : confidence threshold
        self.lbl_conf = QLabel(self)
        self.lbl_conf.setText('65%')
        self.lbl_conf.setFont(font)
        self.lbl_conf.setStyleSheet('background-color: #FFFFFF')
        self.lbl_conf.setFixedHeight(60)
        # self.lbl_conf.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        btn_conf = QPushButton('분류 정확도 임계치', self)
        btn_conf.setFont(font)
        btn_conf.setStatusTip('분류 정확도 임계치를 설정합니다. 해당 값 미만은 인식하지 않습니다. 21% ~ 99%')
        btn_conf.setToolTip('분류 정확도 임계치를 설정합니다. 해당 값 미만은 인식하지 않습니다.\n21% ~ 99%')
        btn_conf.setStyleSheet("background-image: url(./icon/background.jpg);")
        btn_conf.setFixedHeight(60)
        # btn_conf.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        btn_conf.clicked.connect(self.conf)

        # line_thickness : bounding box thickness (pixels)
        self.lbl_ltk = QLabel(self)
        self.lbl_ltk.setNum(2)
        self.lbl_ltk.setFont(font)
        self.lbl_ltk.setStyleSheet('background-color: #FFFFFF')
        self.lbl_ltk.setFixedHeight(60)
        # self.lbl_ltk.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        btn_ltk = QPushButton('박스 굵기', self)
        btn_ltk.setFont(font)
        btn_ltk.setStatusTip('바운딩박스 굵기(픽셀값)을 설정합니다.')
        btn_ltk.setToolTip('바운딩박스 굵기(픽셀값)을 설정합니다.')
        btn_ltk.setStyleSheet("background-image: url(./icon/background.jpg);")
        btn_ltk.setFixedHeight(60)
        # btn_ltk.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        btn_ltk.clicked.connect(self.ltk)

        # self.btn_start = QPushButton('Start', self)
        # self.btn_start.setFont(font)
        # self.btn_start.clicked.connect(self.run)
        # self.btn_stop = QPushButton('실행 중지', self)
        # self.btn_stop.setFont(font)
        # self.btn_stop.setStyleSheet("background-image: url(./icon/background.jpg);")
        # self.btn_stop.setEnabled(False)
        # self.btn_stop.clicked.connect(self.stop)

        self.btn_prev = QPushButton('이전', self)
        self.btn_prev.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 1211 창 크기 변경 시 버튼 크기 자동 조절
        self.btn_prev.setFont(font)
        self.btn_prev.setEnabled(False)
        self.btn_prev.setToolTip('이전 파일\n단축키:A')
        self.btn_prev.setStatusTip('이전 파일 단축키:A')
        self.btn_prev.setStyleSheet("background-image: url(./icon/background.jpg);")
        # self.btn_prev.setStyleSheet("""QPushButton{
        #                 background-image: url('./icon/background.jpg');
        #                 border-style: inset;
        #                 border-width: 1px;
        #                 border-radius:15px;
        #                 padding: 6px;
        #                 }
        #                 QPushButton:pressed {
        #                 background-color: rgb(250, 250, 250);
        #                 border-style: outset;}""")
        self.btn_prev.clicked.connect(self.frame_reset)
        self.btn_prev.clicked.connect(self.prev)

        self.btn_next = QPushButton('다음', self)
        self.btn_next.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 1211 창 크기 변경 시 버튼 크기 자동 조절
        self.btn_next.setFont(font)
        self.btn_next.setEnabled(False)
        self.btn_next.setStyleSheet("background-image: url(./icon/background.jpg);")
        # self.btn_next.setStyleSheet('border-image: url(./icon/next.png);')
        # self.btn_next.setStyleSheet("""QPushButton{
        #                 background-image: url('./icon/background.jpg');
        #                 border-style: inset;
        #                 border-width: 1px;
        #                 border-radius:15px;
        #                 padding: 6px;
        #                 }
        #                 QPushButton:pressed {
        #                 background-color: rgb(250, 250, 250);
        #                 border-style: outset;}""")
        self.btn_next.setToolTip('다음 파일\n단축키:D')
        self.btn_next.setStatusTip('다음 파일 단축키:D')
        self.btn_next.clicked.connect(self.next)

        self.img2 = QLabel()
        self.img2.mouseDoubleClickEvent = self.magnify
        self.img2.setToolTip('더블클릭시 확대')
        self.img2.setStatusTip('더블클릭시 확대')
        self.img2.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        self.img2.setScaledContents(True)
        self.img2.setPixmap(QPixmap('./icon/main.JPG'))
        # self.img2.setFixedSize(int(width * 2 / 3), int(height * 2 / 3))

        self.sliconf = QSlider(Qt.Horizontal, self)
        self.sliconf.setRange(21, 99)
        self.sliconf.setSingleStep(1)
        self.sliconf.setValue(65)
        self.sliconf.setMinimumHeight(50)
        self.sliconf.setStyleSheet("""QSlider::groove:horizontal {
                                    height: 3px;
                                    margin: 0px 20;
                                    background-color: rgb(200, 200, 200);
                                    }
                                    QSlider::handle:horizontal {
                                    border-image: url(./icon/logo3.png);
                                    border: none;
                                    height: 50px;
                                    width: 40px;
                                    margin: -20px -20;
                                    }""")
        self.sliconf.valueChanged.connect(self.conf_chg)

        self.sliframe = QSlider(Qt.Horizontal, self)  # Main tab / set lbl_frame slider
        self.sliframe.setSingleStep(1)
        self.sliframe.setRange(0, 0)
        self.sliframe.setValue(0)
        self.sliframe.setEnabled(False)
        self.sliframe.setMinimumHeight(50)
        # self.sliframe.setStyleSheet("background-image: url(./icon/background.jpg);")QSlider::handle:horizontal {background: green;}QSlider::groove:horizontal {background: blue;}
        # self.sliframe.setStyleSheet("QSlider::groove:horizontal {background: blue;} QSlider::handle:horizontal {background-image: url(./icon/main.ico);height: 1px; width: 18px; }")
        self.sliframe.setStyleSheet("""QSlider::groove:horizontal {
                                    height: 3px;
                                    margin: 0px 20;
                                    background-color: rgb(200, 200, 200);
                                    }
                                    QSlider::handle:horizontal {
                                    border-image: url(./icon/logo3.png);
                                    border: none;
                                    height: 50px;
                                    width: 40px;
                                    margin: -20px -20;
                                    }""")
        self.sliframe.valueChanged.connect(self.frame_chg)

        self.table = QTableWidget()  # Result tab / tableviewer
        if self.save_status:
            self.table.setColumnCount(len(self.app_data['columns']))
            self.table.setHorizontalHeaderLabels(self.app_data['columns'])
        else:
            row, column = self.cache.shape
            self.table.setRowCount(row)
            self.table.setColumnCount(column)
            self.table.setHorizontalHeaderLabels(self.cache.columns)
            for i in range(row):
                for j in range(column):
                    try:
                        self.table.setItem(i, j, QTableWidgetItem(str(self.cache.iloc[i, j])))
                    except Exception as e:
                        logger.warning('writing cache to result table error', exc_info=e)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked)
        self.table.setItemDelegateForColumn(0, ReadOnlyDelegate(self.table))  # 파일명 수정 X
        self.table.itemChanged.connect(self.df_chg)
        self.table.cellDoubleClicked.connect(self.rename2)

        self.sample_table = QTableWidget()  # Main tab / tableviewer
        self.sample_table.setFocusPolicy(Qt.NoFocus)
        self.sample_table.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)  # 1211 창 크기 변경 시 버튼 크기 자동 조절
        self.sample_table.setRowCount(1)
        self.sample_table.setColumnCount(len(self.app_data['columns']))
        self.sample_table.setHorizontalHeaderLabels(self.app_data['columns'])
        # self.sample_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.sample_table.setEditTriggers(QAbstractItemView.DoubleClicked)
        self.sample_table.setItemDelegateForColumn(0, ReadOnlyDelegate(self.sample_table))  # 파일명 수정 X
        self.sample_table.setStatusTip("'테이블 저장' 버튼을 눌러 현재 보이는 데이터를 기록합니다.")
        self.sample_table.setToolTip("'테이블 저장' 버튼을 눌러 현재 보이는 데이터를 기록합니다.")
        # self.sample_table.setDragDropMode(QAbstractItemView.InternalMove)

        self.sample_table.setColumnWidth(0, 70)  # 파일명
        self.sample_table.setColumnWidth(1, 35)  # 년
        for idx in range(2, 6):  # 월,일,시간,분
            self.sample_table.setColumnWidth(idx, 14)
        self.sample_table.setColumnWidth(6, 42)  # 주야간
        self.sample_table.setColumnWidth(7, 70)  # 국명
        self.sample_table.setColumnWidth(8, 70)  # 학명
        self.sample_table.setColumnWidth(9, 42)  # 개체 수
        self.sample_table.setColumnWidth(10, 28)  # 온도
        self.sample_table.setColumnWidth(11, 70)  # 최대정확도
        self.sample_table.setColumnWidth(12, 70)  # 최소정확도

        self.sample_table.horizontalHeader().setStretchLastSection(True)  # 마지막열 크기조절
        self.sample_table.cellDoubleClicked.connect(self.rename)
        self.sample_table.cellChanged.connect(self.set_animal2)
        self.sample_table.setItem(0, 7, QTableWidgetItem('미동정'))

        self.btn_submit = QPushButton('테이블 저장', self)  # Main tab / Save button
        self.btn_submit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 1211 창 크기 변경 시 버튼 크기 자동 조절
        self.btn_submit.setFont(font)
        self.btn_submit.setEnabled(False)
        self.btn_submit.setStatusTip("클릭 시 데이터를 기록합니다. 단축키:S")
        self.btn_submit.setToolTip("클릭 시 데이터를 기록합니다.\n단축키:S")
        self.btn_submit.clicked.connect(self.submit)
        self.btn_submit.setStyleSheet("background-image: url(./icon/background.jpg);")
        # self.btn_submit.setStyleSheet("""QPushButton{
        #                 background-image: url('./icon/background.jpg');
        #                 border-style: inset;
        #                 border-width: 1px;
        #                 border-radius:15px;
        #                 padding: 6px;
        #                 }
        #                 QPushButton:pressed {
        #                 background-color: rgb(250, 250, 250);
        #                 border-style: outset;}""")

        self.btn_export = QPushButton('엑셀로 저장', self)  # Result tab / Data to excel
        self.btn_export.setFont(font)
        self.btn_export.setStatusTip('결과물을 xlsx파일로 저장합니다.')
        self.btn_export.setToolTip('결과물을 xlsx파일로 저장합니다.')
        self.btn_export.clicked.connect(self.save_df)
        self.btn_export.setStyleSheet("background-image: url(./icon/background.jpg);")
        self.btn_export.setFixedHeight(30)

        self.btn_add = QPushButton('칼럼 추가', self)  # Result tab / add column button
        self.btn_add.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 1211 창 크기 변경 시 버튼 크기 자동 조절
        self.btn_add.setFont(font)
        self.btn_add.setStatusTip('테이블의 칼럼을 추가합니다.')
        self.btn_add.setToolTip('테이블의 칼럼을 추가합니다.')
        self.btn_add.clicked.connect(self.add_column)
        self.btn_add.setStyleSheet("background-image: url(./icon/background.jpg);")
        # self.btn_add.setStyleSheet("""QPushButton{
        #                 background-image: url('./icon/background.jpg');
        #                 border-style: inset;
        #                 border-width: 1px;
        #                 border-radius:15px;
        #                 padding: 6px;
        #                 }
        #                 QPushButton:pressed {
        #                 background-color: rgb(250, 250, 250);
        #                 border-style: outset;}""")

        self.btn_remove = QPushButton('칼럼 제거', self)  # Result tab / remove column button
        self.btn_remove.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 1211 창 크기 변경 시 버튼 크기 자동 조절
        self.btn_remove.setFont(font)
        self.btn_remove.setStatusTip('테이블의 칼럼을 제거합니다. 기본값은 제거할 수 없습니다.')
        self.btn_remove.setToolTip('테이블의 칼럼을 제거합니다.\n기본값은 제거할 수 없습니다.')
        self.btn_remove.clicked.connect(self.remove_column)
        self.btn_remove.setStyleSheet("background-image: url(./icon/background.jpg);")
        # self.btn_remove.setStyleSheet("""QPushButton{
        #         background-image: url('./icon/background.jpg');
        #         border-style: inset;
        #         border-width: 1px;
        #         border-radius:15px;
        #         padding: 6px;
        #         }
        #         QPushButton:pressed {
        #         background-color: rgb(250, 250, 250);
        #         border-style: outset;}""")

        self.btn_pp = QPushButton(self)  # Main tab / video play/pause button
        self.btn_pp.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 1211 창 크기 변경 시 버튼 크기 자동 조절
        self.btn_pp.setCheckable(True)
        self.btn_pp.setChecked(False)
        self.btn_pp.setEnabled(False)
        self.btn_pp.setStyleSheet("background-image: url(./icon/background.jpg); ")
        # self.btn_pp.setStyleSheet("""QPushButton{
        #                             border-image: url(./icon/play.png);}
        #                             QPushButton:checked {
        #                             border-image: url(./icon/pause.png);
        #                             }""")
        # self.btn_pp.setStyleSheet("""QPushButton{
        # background-image: url('./icon/background.jpg');
        # border-style: inset;
        # border-width: 1px;
        # border-radius:10px;
        # padding: 6px;
        # }
        # QPushButton:pressed {
        # background-color: rgb(250, 250, 250);
        # border-style: outset;}""")
        icon_pp = QIcon()
        icon_pp.addPixmap(QPixmap("./icon/pause.png"), QIcon.Normal, QIcon.On)
        icon_pp.addPixmap(QPixmap("./icon/play.png"), QIcon.Active, QIcon.Off)
        icon_pp.addPixmap(QPixmap("./icon/pause.png"), QIcon.Active, QIcon.On)
        self.btn_pp.setIcon(icon_pp)
        self.btn_pp.setStatusTip('재생 / 정지 단축키:P')
        self.btn_pp.setToolTip('재생 / 정지\n단축키:P')
        self.btn_pp.clicked.connect(self.pp)

        self.timebox = QSpinBox(self)
        self.timebox.setRange(-1000, 1000)
        self.timebox.setSuffix(" H")
        self.timebox.setValue(0)
        self.timebox.valueChanged.connect(self.edit_time)

        self.lbl_frame = QLabel('1/0', self)  # Main tab / Show lbl_frame
        self.lbl_frame.setAlignment(Qt.AlignCenter)
        self.lbl_frame.setStatusTip('더블클릭시 원하는 프레임으로 이동합니다.')
        self.lbl_frame.setToolTip('더블클릭시 원하는 프레임으로 이동합니다.')
        self.lbl_frame.mouseDoubleClickEvent = self.set_frame

        speed_opt = QComboBox(self)  # Main tab / Speed option
        speed_opt.setFocusPolicy(Qt.NoFocus)
        speed_opt.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 1211 창 크기 변경 시 버튼 크기 자동 조절
        [speed_opt.addItem(i) for i in ["Very Slow", "Slow", "Normal", "Fast"]]
        speed_opt.setFont(font)
        speed_opt.setStatusTip('속도를 조절합니다.')
        speed_opt.setToolTip('속도를 조절합니다')
        speed_opt.setCurrentIndex(2)
        speed_opt.activated[str].connect(self.det_thread.speed_slow)
        speed_opt.setStyleSheet("background-image: url(./icon/background.jpg);")

        self.lbl_cnt = QLabel()  # Status bar / progress
        self.lbl_cnt.setStatusTip('더블클릭시 이미지/영상을 이동합니다.')
        self.lbl_cnt.setToolTip('더블클릭시 이미지/영상을 이동합니다.')
        self.lbl_cnt.setFixedWidth(80)
        self.lbl_cnt.mouseDoubleClickEvent = self.move_file

        self.lbl_info = QLabel()

        btn_rm = QPushButton('데이터 제거', self)
        btn_rm.setFont(font)
        btn_rm.setStatusTip('원하는 데이터(행)을 제거합니다.')
        btn_rm.setToolTip('원하는 데이터(행)을 제거합니다.')
        btn_rm.clicked.connect(self.rm_data)
        btn_rm.setStyleSheet("background-image: url(./icon/background.jpg);")
        btn_rm.setFixedHeight(30)

        self.btn_add_li = QPushButton('간편 목록 추가')
        self.btn_add_li.clicked.connect(self.add_ani)
        self.btn_add_li.setStyleSheet("background-image: url(./icon/background.jpg);")
        # self.btn_add_li.setStyleSheet("""QPushButton{
        #                 background-image: url('./icon/background.jpg');
        #                 border-style: outset;
        #                 border-width: 1px;
        #                 border-radius:15px;
        #                 border-color: rgb(120, 120, 120);
        #                 padding: 6px;}
        #                 QPushButton:pressed {
        #                 border-color: rgb(255, 255, 255);
        #                 border-style: inset;}""")
        self.btn_add_li.setStatusTip('국명을 리스트에 추가합니다.')
        self.btn_add_li.setToolTip('국명을 리스트에 추가합니다.')
        self.btn_add_li.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 1211 창 크기 변경 시 버튼 크기 자동 조절
        self.btn_add_li.setFont(font)
        self.btn_rm_file = QPushButton('파일 제거')
        self.btn_rm_file.clicked.connect(self.erase_file)
        self.btn_rm_file.setEnabled(False)
        self.btn_rm_file.setStyleSheet("background-color: #FF0000")
        # self.btn_rm_file.setStyleSheet("""QPushButton{
        #                 background-image: url('./icon/background.jpg');
        #                 border-style: outset;
        #                 border-width: 1px;
        #                 border-radius:15px;
        #                 border-color: rgb(100, 100, 100);
        #                 padding: 6px;}
        #                 QPushButton:pressed {
        #                 border-color: rgb(255, 255, 255);
        #                 border-style: inset;}""")
        self.btn_rm_file.setStatusTip('현재 파일을 제거합니다. 단축키:X')
        self.btn_rm_file.setToolTip('현재 파일을 제거합니다.\n단축키:X')
        self.btn_rm_file.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 1211 창 크기 변경 시 버튼 크기 자동 조절
        self.btn_rm_file.setFont(font)

        logo = QLabel()
        logo.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        logo.setScaledContents(True)
        bigleader = QPixmap('./icon/bigleader.png')
        bigleader = bigleader.scaled(300, 80, Qt.KeepAspectRatio)
        logo.setPixmap(bigleader)

        logo2 = QLabel()
        logo2.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        logo2.setScaledContents(True)
        bigleader2 = QPixmap('./icon/bigleader.png')
        bigleader2 = bigleader2.scaled(300, 100, Qt.KeepAspectRatio)
        logo2.setPixmap(bigleader2)

        self.chk_hlbl = QCheckBox('동물 이름 숨김')
        self.chk_hlbl.setStatusTip('체크시 처리된 영상이 분류된 동물 이름을 띄우지 않습니다.')
        self.chk_hlbl.setToolTip('체크시 처리된 영상이 분류된 동물 이름을 띄우지 않습니다.')
        self.chk_hlbl.stateChanged.connect(self.hlbl)
        self.chk_hlbl.setFont(font)
        self.chk_hconf = QCheckBox('분류 정확도 숨김')
        self.chk_hconf.setChecked(True)
        self.chk_hconf.setStatusTip('체크시 처리된 영상이 신뢰도값을 띄우지 않습니다.')
        self.chk_hconf.setToolTip('체크시 처리된 영상이 신뢰도값을 띄우지 않습니다.')
        self.chk_hconf.stateChanged.connect(self.hconf)
        self.chk_hconf.setFont(font)
        self.chk_unknown = QCheckBox('Unknown')
        self.chk_unknown.setStatusTip('체크시 신뢰도값이 0.2보다 크고 설정된 신뢰도 하한값 보다 작으면 unknown으로 라벨링 합니다.')
        self.chk_unknown.setToolTip('체크시 신뢰도값이 0.2보다 크고 설정된 신뢰도 하한값 보다 작으면 unknown으로 라벨링 합니다.')
        self.chk_unknown.stateChanged.connect(self.unknown)
        self.chk_unknown.setFont(font)
        self.chk_device = QCheckBox('CPU 사용')
        self.chk_device.setStatusTip('체크시 딥러닝 장치로 CPU를 사용합니다.')
        self.chk_device.setToolTip('체크시 딥러닝 장치로 CPU를 사용합니다.')
        self.chk_device.setFont(font)
        self.chk_autorename = QCheckBox('자동 이름 변경')
        self.chk_autorename.setStatusTip('체크시 테이블 저장시에 자동으로 파일이름을 "현재파일이름_동물명" 으로 변경합니다.')
        self.chk_autorename.setToolTip('체크시 테이블 저장시에 자동으로 파일이름을 "현재파일이름_동물명" 으로 변경합니다.')
        self.chk_autorename.setFont(font)

        self.btn_clear = QPushButton('초기화')
        self.btn_clear.setFont(font)
        self.btn_clear.setStyleSheet("background-image: url(./icon/background.jpg);")
        self.btn_clear.setStatusTip('테이블을 초기화합니다.')
        self.btn_clear.setToolTip('테이블을 초기화합니다.')
        self.btn_clear.clicked.connect(self.clear)
        self.btn_clear.setFixedHeight(30)

        self.btn_search = QPushButton('파일번호 검색')
        self.btn_search.setFont(font)
        self.btn_search.setStyleSheet("background-image: url(./icon/background.jpg);")
        self.btn_search.setStatusTip('파일이름이 몇 번째 파일인지 검색합니다.')
        self.btn_search.setToolTip('파일이름이 몇 번째 파일인지 검색합니다.')
        self.btn_search.clicked.connect(self.search)
        self.btn_search.setFixedHeight(30)

        self.list_animal = QListWidget()
        self.list_animal.setFont(big_font)
        self.list_animal.setFocusPolicy(Qt.NoFocus)
        self.animal_none = self.list_animal.item(0)
        self.animal_wildboar = self.list_animal.item(1)
        self.animal_waterdeer = self.list_animal.item(2)
        self.animal_goral = self.list_animal.item(3)
        for i, cat in enumerate(self.app_data['category']):
            self.list_animal.addItem(cat)
            if cat == DEFAULT_ANIMAL[0]:
                self.animal_none = self.list_animal.item(i)
            if cat == DEFAULT_ANIMAL[1]:
                self.animal_wildboar = self.list_animal.item(i)
            if cat == DEFAULT_ANIMAL[2]:
                self.animal_waterdeer = self.list_animal.item(i)
            if cat == DEFAULT_ANIMAL[3]:
                self.animal_goral = self.list_animal.item(i)
        self.list_animal.setCurrentItem(self.animal_none)
        self.list_animal.setDragDropMode(QAbstractItemView.InternalMove)
        self.list_animal.setDragEnabled(False)
        self.list_animal.currentItemChanged.connect(self.set_animal)  # 클릭 시
        self.list_animal.itemDoubleClicked.connect(self.rm_ani)  # 더블클릭 시

        icon_lock = QIcon()
        icon_lock.addPixmap(QPixmap("./icon/lock.PNG"), QIcon.Active, QIcon.Off)
        icon_lock.addPixmap(QPixmap("./icon/unlock.PNG"), QIcon.Active, QIcon.On)
        self.btn_animal_lock = QPushButton()
        self.btn_animal_lock.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 1211 창 크기 변경 시 버튼 크기 자동 조절
        self.btn_animal_lock.setCheckable(True)
        self.btn_animal_lock.setChecked(False)
        self.btn_animal_lock.setIcon(icon_lock)
        self.btn_animal_lock.clicked.connect(self.animal_lock)
        self.btn_animal_lock.setStyleSheet("background-color:#FFFFFF")
        self.btn_animal_lock.setToolTip("체크시 동물목록을 드래그&드롭으로 위치변경 및 삭제 가능")
        self.btn_animal_lock.setStatusTip("체크시 동물목록을 드래그&드롭으로 위치변경 및 삭제 가능")

        # girdmain
        gridmain.addWidget(self.img2, 0, 0, 80, 76)
        gridmain.addWidget(logo, 0, 76, 10, 24)
        gridmain.addWidget(self.btn_add_li, 10, 76, 5, 10)
        gridmain.addWidget(self.btn_animal_lock, 10, 96, 5, 4)
        gridmain.addWidget(self.list_animal, 15, 76, 70, 24)  # 동물그룹
        gridmain.addWidget(self.timebox, 80, 0, 5, 2)
        gridmain.addWidget(self.lbl_frame, 80, 2, 5, 6)
        gridmain.addWidget(self.sliframe, 80, 8, 5, 60)
        gridmain.addWidget(self.btn_submit, 85, 84, 7, 8)  # 테이블 저장
        gridmain.addWidget(self.btn_rm_file, 85, 92, 7, 8)
        gridmain.addWidget(speed_opt, 80, 68, 5, 8)  # 속도조절
        gridmain.addWidget(self.noon_night(), 85, 76, 7, 8)  # 주간/야간 버튼
        gridmain.addWidget(self.btn_prev, 92, 76, 7, 8)  # 이전
        gridmain.addWidget(self.btn_next, 92, 92, 7, 8)  # 다음
        gridmain.addWidget(self.btn_pp, 92, 84, 7, 8)  # 재생/정지 버튼
        gridmain.addWidget(self.sample_table, 85, 0, 14, 60)  #
        gridmain.addWidget(self.btn_add, 85, 60, 7, 8)  # 1211
        gridmain.addWidget(self.btn_remove, 92, 60, 7, 8)  #
        gridmain.addWidget(self.temperature(), 85, 68, 14, 8)  # 온도

        # gridopt
        gridopt.addWidget(self.btn_source, 0, 0, 1, 5)
        gridopt.addWidget(self.lbl_source, 0, 5, 1, 10)
        gridopt.addWidget(btn_conf, 1, 0, 1, 5)
        gridopt.addWidget(self.lbl_conf, 1, 5, 1, 3)
        gridopt.addWidget(self.sliconf, 1, 8, 1, 12)
        gridopt.addWidget(btn_ltk, 2, 0, 1, 5)
        gridopt.addWidget(self.lbl_ltk, 2, 5, 1, 2)
        gridopt.addWidget(self.chk_hlbl, 4, 0, 1, 5)
        gridopt.addWidget(self.chk_hconf, 5, 0, 1, 5)
        gridopt.addWidget(self.chk_unknown, 6, 0, 1, 5)
        gridopt.addWidget(self.imgsz(), 3, 0, 1, 5)
        gridopt.addWidget(self.weights(), 3, 7, 1, 5)
        gridopt.addWidget(self.autogroup(), 4, 7, 3, 5)
        gridopt.addWidget(self.chk_autorename, 3, 14, 1, 4)

        # gridtable
        gridtable.addWidget(self.btn_export, 5, 3, 1, 1)
        gridtable.addWidget(btn_rm, 5, 2, 1, 1)
        gridtable.addWidget(self.btn_clear, 5, 0, 1, 1)
        gridtable.addWidget(self.btn_search, 5, 1, 1, 1)
        gridtable.addWidget(self.table, 0, 0, 5, 4)

        main = QWidget()
        main.setLayout(gridmain)
        main.setFocusPolicy(Qt.NoFocus)

        self.options = QWidget()
        self.options.setLayout(gridopt)

        table = QWidget()
        table.setLayout(gridtable)

        self.tabs = QTabWidget(self)
        self.tabs.setFocusPolicy(Qt.NoFocus)
        self.tabs.addTab(main, 'Main')
        self.tabs.addTab(self.options, 'Options')
        self.tabs.addTab(table, 'Result')

    def prev(self):
        self.det_thread.prev()

    def next(self):
        self.btn_next.setEnabled(False)
        self.det_thread.next()

    def animal_lock(self):  # animal_list drag enable
        if self.btn_animal_lock.isChecked():
            self.list_animal.setDragEnabled(True)
            QMessageBox.information(self, '확인', '편집 종료 후 다시 눌러 잠그셔야 변경내용이 저장됩니다.')
        else:
            self.btn_animal_lock.setEnabled(False)
            self.list_animal.setDragEnabled(False)
            self.app_data['category'] = [self.list_animal.item(i).text() for i in range(self.list_animal.count())]
            with open(self.app_data_dir, 'w', encoding='utf-8') as f:
                yaml.dump(self.app_data, f)
            QTest.qWait(1000)
            self.btn_animal_lock.setEnabled(True)

    def search(self):  # search current file index  번호 검색
        if self.det_thread.isRunning():
            dlg = QInputDialog(self)
            dlg.setWindowIcon(icon)
            dlg.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint)
            dlg.setInputMode(QInputDialog.TextInput)
            dlg.setWindowTitle('번호 검색')
            dlg.setLabelText('검색할 파일명을 입력하세요.\n* 확장자도 입력해주세요\n* 대소문자에 주의해주세요')
            dlg.setTextValue('Image.JPG')
            dlg.resize(700, 100)
            ok = dlg.exec_()
            text = dlg.textValue()
            if ok and text:
                try:
                    index = self.det_thread.files.index(self.det_thread.source.replace('/', '\\') + '\\' + text)
                except ValueError:
                    QMessageBox.warning(self, 'File not found!',
                                        f'{text}파일을 찾을 수 없습니다.\n번호 검색은 확장자가 정확해야하며 대소문자를 구분합니다.')
                else:
                    QMessageBox.information(self, '번호 검색', f'{text}파일은 {index + 1}번째 파일입니다.')

    def clear(self):  # reset result table  result테이블 초기화
        if not self.save_status:
            QMessageBox.warning(self, '저장되지 않은 데이터', '현재 테이블은 저장되지 않았습니다')
        reply = QMessageBox.question(self, '테이블 초기화', '현재 테이블을 초기화하시겠습니까?', QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.table.clearContents()
            self.table.setRowCount(0)
            self.cache = self.cache[0:0]
            self.save_status = True

    def erase_file(self, e):  # remove image & label at trash
        file = Path(self.det_thread.path)
        reply = QMessageBox.question(self, '파일 제거', f"다음 파일을 제거하시겠습니까?\n{file}",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            try:
                if self.det_thread.video_flag:
                    self.det_thread.vid_cap.release()
                shutil.move(Path(file), Path(file).parent / 'trash')
            except FileNotFoundError:
                QMessageBox.critical(self, '파일 없음', '파일을 찾을 수 없습니다.')
            except (PermissionError, OSError) as e:
                print(e)
                QMessageBox.critical(self, '사용중인 파일', '파일이 다른곳에서 사용중이어서 삭제가 불가능합니다.')
            else:
                self.det_thread.next()

    def rename(self, _, col):  # rename file name at main tab  메인 탭에서 파일명을 변경합니다.
        if self.det_thread.isRunning():
            if col == 0:
                if self.det_thread.mode == 'image':
                    file = self.det_thread.path
                    suffix = file.split('.')[-1]
                    fname = file.split("\\")[-1].split(".")[0]
                    file_len = len(file.split("\\")[-1])
                    path = file[:-file_len]
                    dlg = QInputDialog(self)
                    dlg.setWindowIcon(icon)
                    dlg.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint)
                    dlg.setInputMode(QInputDialog.TextInput)
                    dlg.setWindowTitle('파일명 변경')
                    dlg.setLabelText(f'변경할 파일명을 입력하세요.\n현재파일명:{fname}')
                    dlg.setTextValue(f'{fname}')
                    dlg.resize(700, 100)
                    ok = dlg.exec_()
                    text = re.sub(r"[^a-zA-Z0-9가-힣_ \-(){}\[\]]", "", dlg.textValue().strip()) + '.' + suffix
                    if ok and text:
                        new = path + text
                        try:
                            os.rename(file, new)
                        except (PermissionError, FileNotFoundError):
                            QMessageBox.critical(self, '변경 실패!',
                                                 '파일명 변경에 실패했습니다.\n해당 파일이 열려있거나 혹은 같은 이름의 파일이 존재합니다.')
                        else:
                            self.det_thread.files[self.det_thread.count] = new
                            self.det_thread.path = path + text
                            self.sample_table.setItem(0, 0, QTableWidgetItem(text))
                            QMessageBox.information(self, '변경 완료!',
                                                    f'파일명이 성공적으로 변경되었습니다.\n{fname + suffix} -> {text}')
                else:
                    QMessageBox.critical(self, '변경 실패!', '비디오는 이곳에서 파일명을 변경할 수 없습니다.\nResult 탭에서 시도해주세요!')

    def rename2(self, row, col):  # rename file name at result tab  결과탭에서 파일명을 변경합니다.
        if self.det_thread.isRunning():
            if col == 0:
                path = self.det_thread.source.replace('/', '\\')
                file = self.table.item(row, col).text()
                old = path + '\\' + file
                fname, suffix = file.split('.')
                dlg = QInputDialog(self)
                dlg.setWindowIcon(icon)
                dlg.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint)
                dlg.setInputMode(QInputDialog.TextInput)
                dlg.setWindowTitle('파일명 변경')
                dlg.setLabelText(f'변경할 파일명을 입력하세요.\n현재파일명:{fname}')
                dlg.setTextValue(f'{fname}')
                dlg.resize(700, 100)
                ok = dlg.exec_()
                text = re.sub(r"[^a-zA-Z0-9가-힣_ \-(){}\[\]]", "", dlg.textValue().strip()) + '.' + suffix
                if ok and text:
                    new = path + '\\' + text
                    try:
                        os.rename(old, new)
                    except (OSError, PermissionError, FileNotFoundError):
                        QMessageBox.critical(self, '변경 실패!',
                                             '파일명 변경에 실패했습니다.\n같은 이름의 파일이 존재하거나, 해당 파일이 열려있거나, Main탭에서 사용중입니다.')
                    else:
                        self.det_thread.files[self.det_thread.files.index(old)] = new
                        self.table.setItem(row, col, QTableWidgetItem(text))
                        self.save_status = False
                        QMessageBox.information(self, '변경 완료!', f'파일명이 성공적으로 변경되었습니다.\n{file} -> {text}')

    def background(self, path):  # auto mode background path
        self.background_list.append(str(Path(path)).split('/')[-1])

    def detect(self, path):  # auto mode detect path
        self.detect_list.append(str(Path(path)).split('/')[-1])

    def auto_finish(self, e):  # auto mode finish
        if e:
            self.lbl_cnt.setText('파일이동중')
            p = Path(self.det_thread.source)
            QTest.qWait(1000)
            self.det_thread.vid_cap.release()
            self.det_thread.terminate()
            os.makedirs(p / 'Background', exist_ok=True)
            for file in self.background_list:
                try:
                    shutil.move(str(p / file), str(p / 'Background'))
                except (FileNotFoundError, PermissionError):
                    pass
            os.makedirs(p / 'Detect', exist_ok=True)
            for file in self.detect_list:
                try:
                    shutil.move(str(p / file), str(p / 'Detect'))
                except (FileNotFoundError, PermissionError):
                    pass
            self.btn_submit.setEnabled(False)
            self.btn_pp.setChecked(False)
            self.btn_pp.setEnabled(False)
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(False)
            self.sliframe.setEnabled(False)
            self.img2.setPixmap(QPixmap('./icon/main.JPG'))
            self.lbl_frame.setText('1/0')
            self.sliframe.setValue(1)
            self.lbl_cnt.setText('')
            self.det_thread.status = True
            self.options.setEnabled(True)
            self.btn_source.setEnabled(True)
            self.btn_add.setEnabled(True)
            self.btn_remove.setEnabled(True)
            self.btn_add_li.setEnabled(True)
            self.btn_rm_file.setEnabled(True)
            self.auto.setEnabled(True)
            self.manual.setEnabled(True)
            self.manual.setChecked(True)
            self.btn_weights.setEnabled(True)
            self.chk_device.setEnabled(True)
            self.lbl_info.setText('')
            self.det_thread.auto = False
            self.background_list.clear()
            self.detect_list.clear()
            reply = QMessageBox.question(self, '자동분리완료', '영상분리가 완료되었습니다.\n바로 감지된 사진/영상(Detect폴더)에 대해 수동검수를 시작하시겠습니까?',
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                fname = str(p / 'Detect')
                p = str(Path(fname).resolve())
                files = natsorted(glob(os.path.join(p, '*.*')))  # dir
                images = [x for x in files if x.split('.')[-1].lower() in IMG_FORMATS]
                videos = [x for x in files if x.split('.')[-1].lower() in VID_FORMATS]
                ni, nv = len(images), len(videos)
                nf = ni + nv
                if not nf:
                    QMessageBox.critical(self, '파일 없음', f'해당 폴더에 사진이나 영상을 찾지 못했습니다.\n{p}\n지원되는 확장자는:\n'
                                                        f'사진: {IMG_FORMATS}\n영상: {VID_FORMATS}', QMessageBox.Ok)
                else:
                    self.lbl_source.setText(str(fname))
                    self.det_thread.source = str(fname)
                    self.run()

    def autogroup(self):  # select mode
        gbx = QGroupBox('모드 변경')
        self.auto = QRadioButton('자동분류(Beta)')
        self.auto.clicked.connect(self.automatic)
        self.auto.setToolTip('딥러닝모델이 멧돼지, 고라니, 산양이 감지되었는지 여부를 판단하여 일차분류합니다.')
        self.auto.setStatusTip('딥러닝모델이 멧돼지 고라니, 산양이 감지되었는지 여부를 판단하여 일차분류합니다.')
        self.manual = QRadioButton('수동검수')
        self.manual.setChecked(True)
        self.manual.clicked.connect(self.automatic)
        hbox = QHBoxLayout()
        hbox.addWidget(self.auto)
        hbox.addWidget(self.manual)
        gbx.setLayout(hbox)
        return gbx

    def automatic(self):  # when try to mode change
        if self.det_thread.isRunning():  # 쓰레드가 실행중이면
            reply = QMessageBox.question(self, '자동분류', '검수도중에는 변경할 수 없습니다.\n현재 진행중인 검수를 종료하시겠습니까?',
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:  # 검수 종료
                self.stop()
            else:  # 아니면
                if self.det_thread.auto:
                    self.auto.setChecked(True)
                else:
                    self.manual.setChecked(True)
        else:
            if self.auto.isChecked():
                self.det_thread.auto = True
            else:
                self.det_thread.auto = False

    def autocheck(self, ani_num):  # 0,wild / 1,deer / 2,goral
        if ani_num == 0:  # 멧돼지
            self.list_animal.setCurrentItem(self.animal_wildboar)
        elif ani_num == 1:  # 고라니
            self.list_animal.setCurrentItem(self.animal_waterdeer)
        elif ani_num == 2:
            if len(self.det_thread.names) == 3:
                self.list_animal.setCurrentItem(self.animal_none)  # unknown도 미동정으로
            elif len(self.det_thread.names) == 4:
                self.list_animal.setCurrentItem(self.animal_goral)  # 산양
        elif ani_num == 3:  # unknown
            self.list_animal.setCurrentItem(self.animal_none)  # unknown도 미동정으로
        else:
            self.list_animal.setCurrentItem(self.animal_none)  # 안잡히면 미동정

    def move_file(self, _):  # When file_num doubleclick at main tab (status bar)
        if not self.det_thread.status:
            dlg = QInputDialog(self)
            dlg.setWindowIcon(icon)
            dlg.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint)
            dlg.setInputMode(QInputDialog.IntInput)
            dlg.setWindowTitle('파일 이동')
            dlg.setLabelText("원하는 파일로 이동합니다.")
            dlg.resize(500, 100)
            dlg.setIntRange(1, self.det_thread.nf)
            dlg.setIntValue(self.det_thread.count + 1)
            ok = dlg.exec_()
            num = dlg.intValue()
            if ok:
                self.det_thread.move(num - 1)

    def set_animal(self, _):  # When animal_list clicked
        ani = self.list_animal.currentItem().text()
        self.sample_table.setItem(0, 7, QTableWidgetItem(ani))
        try:
            self.sample_table.setItem(0, 8, QTableWidgetItem(self.animal_db[ani]))
        except KeyError:
            self.sample_table.setItem(0, 8, QTableWidgetItem(''))

    def set_animal2(self, _, col):  # 직접 동물명을 입력할 때 자동으로 학명입력
        if col == 7:
            try:
                self.sample_table.setItem(0, 8, QTableWidgetItem(self.animal_db[self.sample_table.item(0, 7).text()]))
            except KeyError:
                pass

    def add_ani(self):  # 동물 목록 추가
        dlg = QInputDialog(self)
        dlg.setWindowIcon(icon)
        dlg.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint)
        dlg.setInputMode(QInputDialog.TextInput)
        dlg.setWindowTitle('간편 목록 추가')
        dlg.setLabelText("추가할 동물종명(국명)을 입력하세요\n예시) 멧돼지")
        dlg.setTextValue('멧돼지')
        dlg.resize(700, 100)
        ok = dlg.exec_()
        text = dlg.textValue().strip()
        if ok:
            if text in self.app_data['category']:
                QMessageBox.critical(self, 'Error', '이미 추가된 동물종명입니다.')
            elif text not in self.animal_db.keys():
                self.list_animal.addItem(text)
                QMessageBox.warning(self, '동물정보 찾지 못함', f'"{text}" 동물은 학명정보가 등록되지 않았습니다.')
            else:
                self.list_animal.addItem(text)
                self.app_data['category'].append(text)
                with open(self.app_data_dir, 'w', encoding='utf-8') as f:
                    yaml.dump(self.app_data, f)

    def rm_ani(self):  # 동물 목록 제거
        if self.btn_animal_lock.isChecked():
            if self.list_animal.currentItem().text() in DEFAULT_ANIMAL:
                QMessageBox.critical(self, 'Error', '자동인식 동물종은 제거할 수 없습니다.')
            else:
                reply = QMessageBox.warning(self, '목록 제거', f'"{self.list_animal.currentItem().text()}"을(를) 목록에서 제거합니다.',
                                            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                if reply == QMessageBox.Yes:
                    ani = self.list_animal.takeItem(self.list_animal.currentRow()).text()
                    self.app_data['category'].remove(ani)
                    with open(self.app_data_dir, 'w', encoding='utf-8') as f:
                        yaml.dump(self.app_data, f)

    def noon_night(self):  # Main tab / 주야 구분
        self.noon_btn = QRadioButton('주간')
        self.noon_btn.toggled.connect(self.set_noon)
        self.night_btn = QRadioButton('야간')
        self.night_btn.toggled.connect(self.set_night)

        hbox = QHBoxLayout()
        hbox.addWidget(self.noon_btn)
        hbox.addWidget(self.night_btn)

        blackbox = QGroupBox()
        blackbox.setToolTip('컬러사진이면 주간, 흑백사진이면 야간으로 자동판별합니다.\n잘못 판단하면 수동으로 바꿀 수 있습니다.')
        blackbox.setStatusTip('컬러사진이면 주간, 흑백사진이면 야간으로 자동판별합니다. 잘못 판단하면 수동으로 바꿀 수 있습니다.')
        blackbox.setLayout(hbox)
        return blackbox

    def set_noon(self):  # 주간
        self.sample_table.setItem(0, 6, QTableWidgetItem('주간'))

    def set_night(self):  # 야간
        self.sample_table.setItem(0, 6, QTableWidgetItem('야간'))

    def temperature(self):  # 온도
        gbx = QGroupBox()
        gbx.setToolTip('온도를 기록합니다. 화씨는 자동으로 섭씨로 계산됩니다.')
        gbx.setStatusTip('온도를 기록합니다. 화씨는 자동으로 섭씨로 계산됩니다.')
        self.celsius = QRadioButton('섭씨')
        self.celsius.setChecked(True)
        self.celsius.toggled.connect(self.show_temperature)
        self.fahrenheit = QRadioButton('화씨')
        self.temp = QSpinBox(self)
        self.temp.setRange(-1000, 1000)
        self.temp.setSuffix(" °C")
        self.temp.setValue(0)
        self.temp.valueChanged.connect(self.show_temperature)
        hbox = QHBoxLayout()
        hbox.addWidget(self.celsius)
        hbox.addWidget(self.fahrenheit)
        vbox = QVBoxLayout()
        vbox.addLayout(hbox)
        vbox.addWidget(self.temp)
        gbx.setLayout(vbox)
        return gbx

    def show_temperature(self):  # temperature to sample_table
        if self.celsius.isChecked():
            self.temp.setSuffix(" °C")
            self.sample_table.setItem(0, 10, QTableWidgetItem(str(self.temp.value())))
        else:
            self.temp.setSuffix(" °F")
            self.sample_table.setItem(0, 10, QTableWidgetItem(str(round((self.temp.value() - 32) * 5 / 9))))

    def submit(self):  # Main tab / Save and run thread
        self.btn_submit.setEnabled(False)
        if self.chk_autorename.isChecked():
            file = self.det_thread.path
            suffix = file.split('.')[-1]
            fname = file.split("\\")[-1].split(".")[0]
            file_len = len(file.split("\\")[-1])
            path = file[:-file_len]
            ani = self.sample_table.item(0, 7)
            text = fname + f'_{"미동정" if ani is None else ani.text()}' + '.' + suffix
            new = path + text
            try:
                os.rename(file, new)
            except (PermissionError, FileNotFoundError, OSError):
                logger.warning('submit auto rename error')
                QMessageBox.critical(self, '변경 실패!',
                                     '파일명 자동 변경에 실패했습니다.\n해당 파일이 열려있거나 같은 이름의 파일이 존재합니다.')
            else:
                self.det_thread.files[self.det_thread.count] = new
                self.det_thread.path = path + text
                self.sample_table.setItem(0, 0, QTableWidgetItem(text))
        col = self.sample_table.columnCount()
        row = self.table.rowCount()
        self.table.setRowCount(row + 1)
        for i in range(col):
            item = '' if self.sample_table.item(0, i) is None else self.sample_table.item(0, i).text()
            self.table.setItem(row, i, QTableWidgetItem(item))
        self.save_status = False
        QTest.qWait(100)
        try:
            self.cache.to_csv(self.cache_dir, index=False)
        except Exception as e:
            logger.warning('save cache file error "def submit"', exc_info=e)
            QMessageBox.warning(self, '저장 실패', f'자동저장에 실패했습니다.\n{e}')
        self.det_thread.next()

    def cnt(self, s):  # Status bar / set progress
        self.lbl_cnt.setText(s)
        self.sample_table.setItem(0, 10, QTableWidgetItem(''))

    def hlbl(self):  # hide label
        if self.chk_hlbl.isChecked():
            self.det_thread.hide_labels = True
        else:
            self.det_thread.hide_labels = False
            self.chk_hconf.setChecked(False)

    def hconf(self):  # hide conf
        if self.chk_hconf.isChecked():
            self.det_thread.hide_conf = True
        else:
            self.det_thread.hide_conf = False

    def unknown(self):  # show unknown
        if self.chk_unknown.isChecked():
            self.det_thread.unknown = True
        else:
            self.det_thread.unknown = False

    def weights(self):  # select model
        gbx = QGroupBox('모델')
        self.btn_weights = QComboBox(self)
        [self.btn_weights.addItem(i) for i in ["멧돼지와 고라니", "멧돼지와 고라니와 산양", "사용 안함"]]
        self.btn_weights.setFixedHeight(60)
        self.btn_weights.setStyleSheet("background-image: url(./icon/background.jpg);")
        self.btn_weights.setToolTip('어떤 모델을 사용하여 영상처리를 진행할지 선택합니다.')
        self.btn_weights.setStatusTip('어떤 모델을 사용하여 영상처리를 진행할지 선택합니다.')
        self.btn_weights.activated[str].connect(self.det_thread.model)
        self.btn_weights.activated[str].connect(self.weight_select)
        gbox = QGridLayout()
        gbox.addWidget(self.btn_weights, 0, 0, 1, 3)
        gbox.addWidget(self.chk_device, 0, 3, 1, 1)
        gbx.setLayout(gbox)
        gbx.setFixedHeight(100)
        return gbx

    def weight_select(self, menu):
        if menu == '사용 안함':
            self.manual.click()
            self.manual.setEnabled(False)
            self.auto.setEnabled(False)
        else:
            self.manual.setEnabled(True)
            self.auto.setEnabled(True)

    def source(self):  # 폴더 선택
        if not self.det_thread.isRunning():
            QMessageBox.information(self, '주의사항', '프로그램을 사용하는 도중에 해당 폴더 안의 사진 또는 영상의 이름을 변경하거나 삭제하지 마십시오.')
            fname = QFileDialog.getExistingDirectory(self)  # fname = /
            if fname:
                reply = QMessageBox.Yes
                for drive in DRIVE:
                    if fname.startswith(drive):
                        reply = QMessageBox.warning(self, '이동식디스크',
                                                    '이동식디스크(USB, SD카드 등)로 판단됩니다. 파일 삭제 시 복구가 불가능합니다.\n'
                                                    '하드디스크로 복사해서 처리하는 것을 권장합니다.\n계속하시겠습니까?',
                                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.Yes:
                    p = str(Path(str(fname)).resolve())  # os-agnostic absolute path p = \
                    files = natsorted(glob(os.path.join(p, '*.*')))  # dir
                    images = [x for x in files if x.split('.')[-1].lower() in IMG_FORMATS]
                    videos = [x for x in files if x.split('.')[-1].lower() in VID_FORMATS]
                    ni, nv = len(images), len(videos)
                    nf = ni + nv
                    if not nf:
                        QMessageBox.critical(self, '파일 없음', f'해당 폴더에 사진이나 영상을 찾지 못했습니다.\n{p}\n지원되는 확장자는:\n'
                                                            f'사진: {IMG_FORMATS}\n영상: {VID_FORMATS}', QMessageBox.Ok)
                    else:
                        self.lbl_source.setText(str(fname))
                        self.det_thread.source = str(fname)
                        reply = QMessageBox.question(self, '실행',
                                                     f'{str(fname)} 폴더 안의 파일들을 실행하시겠습니까?\n사진 : {ni}개, 영상 : {nv}개',
                                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                        if reply == QMessageBox.Yes:
                            os.makedirs(Path(p) / 'trash', exist_ok=True)
                            self.run()

    def disable(self, e):  # disable btn
        if e:
            self.lbl_source.setEnabled(False)
            self.btn_source.setEnabled(False)
            self.auto.setEnabled(False)
            self.manual.setEnabled(False)
            self.options.setEnabled(True)
            self.btn_weights.setEnabled(False)
            self.chk_device.setEnabled(False)

    def imgsz(self):  # inference size(pixel)
        gbx = QGroupBox('사진 분할개수')
        btn_imgsz = QComboBox(self)
        [btn_imgsz.addItem(i) for i in ["1280", "960", "640", "480", "320", "128"]]
        btn_imgsz.setCurrentIndex(3)
        btn_imgsz.setFixedHeight(60)
        btn_imgsz.setStyleSheet("background-image: url(./icon/background.jpg);")
        btn_imgsz.setToolTip('딥러닝 모델이 분석하기 위한 사진조각개수입니다.\n조각수가 많아질수록 처리시간이 증가하지만 반드시 인식성능이 올라가지는 않습니다.')
        btn_imgsz.setStatusTip('딥러닝 모델이 분석하기 위한 사진조각개수입니다. 조각수가 많아질수록 처리시간이 증가하지만 반드시 인식성능이 올라가지는 않습니다.')
        btn_imgsz.activated[str].connect(self.det_thread.imgsz_opt)
        hbox = QHBoxLayout()
        hbox.addWidget(btn_imgsz)
        gbx.setLayout(hbox)
        gbx.setFixedHeight(100)
        return gbx

    def conf(self):  # conf-thres
        dlg = QInputDialog(self)
        dlg.setWindowIcon(icon)
        dlg.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint)
        dlg.setInputMode(QInputDialog.IntInput)
        dlg.setWindowTitle('분류 정확도 임계치')
        dlg.setLabelText("분류 정확도 임계치를 설정합니다.\n21% ~ 99%")
        dlg.resize(500, 100)
        dlg.setIntRange(20, 99)
        dlg.setIntValue(int(self.lbl_conf.text()[:-1]))
        ok = dlg.exec_()
        num = dlg.intValue()
        if ok:
            self.lbl_conf.setText(str(num) + '%')
            self.sliconf.setValue(num)
            self.det_thread.conf_thres = num / 100

    def conf_chg(self):
        self.lbl_conf.setText(str(self.sliconf.value()) + '%')
        self.det_thread.conf_thres = int(self.sliconf.value()) / 100

    def show_conf(self, conf):
        max_conf, min_conf = conf
        self.sample_table.setItem(0, 11, QTableWidgetItem(str(max_conf)))
        self.sample_table.setItem(0, 12, QTableWidgetItem(str(min_conf)))

    def ltk(self):  # line thickness
        dlg = QInputDialog(self)
        dlg.setWindowIcon(icon)
        dlg.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint)
        dlg.setInputMode(QInputDialog.IntInput)
        dlg.setWindowTitle('박스 굵기')
        dlg.setLabelText("바운딩 박스 및 라벨의 크기(픽셀값)\n1~10")
        dlg.resize(500, 100)
        dlg.setIntRange(1, 10)
        dlg.setIntValue(int(self.lbl_ltk.text()))
        ok = dlg.exec_()
        num = dlg.intValue()
        if ok:
            self.lbl_ltk.setNum(num)
            self.det_thread.line_thickness = num

    def mag(self, raw):  # 확대를 위한 원본 이미지 함수
        self.raw = raw

    def magnify(self, e):  # 확대 이미지 제공
        if not self.det_thread.status:
            cv2.namedWindow('Magnified Image', cv2.WND_PROP_FULLSCREEN)
            cv2.setWindowProperty('Magnified Image', cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
            img = cv2.resize(self.raw, dsize=(width, height), interpolation=cv2.INTER_AREA)
            cv2.imshow('Magnified Image', img)
            if cv2.waitKey() == 27:
                cv2.destroyAllWindows()

    def run(self):  # 쓰레드 실행
        if self.chk_device.isChecked():
            self.det_thread.device = 'cpu'
        else:
            self.det_thread.device = ''
        self.det_thread.start()

    def stop(self):  # 쓰레드 정지
        reply = QMessageBox.warning(self, '프로세스 종료', '현재 프로세스를 종료하고 다른 폴더의 파일을 실행하시겠습니까?\n저장하지 않은 정보는 손실될 수 있습니다.',
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.det_thread.status = False
            QTest.qWait(1000)
            self.det_thread.terminate()
            self.btn_submit.setEnabled(False)
            self.btn_pp.setChecked(False)
            self.btn_pp.setEnabled(False)
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(False)
            self.sliframe.setEnabled(False)
            self.list_animal.setCurrentItem(self.animal_none)
            self.img2.setPixmap(QPixmap('./icon/main.JPG'))
            self.lbl_frame.setText('1/0')
            self.sliframe.setValue(1)
            self.lbl_cnt.setText('')
            self.lbl_info.setText('')
            self.det_thread.status = True
            self.options.setEnabled(True)
            self.btn_source.setEnabled(True)
            self.btn_add.setEnabled(True)
            self.btn_remove.setEnabled(True)
            self.btn_add_li.setEnabled(True)
            self.btn_rm_file.setEnabled(True)
            self.btn_weights.setEnabled(True)
            self.chk_device.setEnabled(True)
            self.sample_table.clearContents()
            if self.auto.isChecked():
                self.det_thread.auto = True
            else:
                self.det_thread.auto = False
            if not self.btn_weights.currentText() == '사용 안함':
                self.auto.setEnabled(True)
                self.manual.setEnabled(True)
        else:
            if self.det_thread.auto:
                self.auto.setChecked(True)
            else:
                self.manual.setChecked(True)

    def df_chg(self, item):  # edit result table change cache data  결과테이블에서 변경한 데이터를 캐시 데이터에 반영
        try:
            self.cache.iloc[item.row(), item.column()] = item.text()
        except IndexError:  # rewrite cache
            col_count = self.table.columnCount()
            row_count = self.table.rowCount()
            headers = [str(self.table.horizontalHeaderItem(i).text()) for i in range(col_count)]
            df_list = []
            for row in range(row_count):
                df_list2 = []
                for col in range(col_count):
                    table_item = '' if self.table.item(row, col) is None else self.table.item(row, col).text()
                    df_list2.append('' if table_item == 'nan' else table_item)
                df_list.append(df_list2)
            self.cache = pd.DataFrame(df_list, columns=headers)
            try:
                self.cache.to_csv(self.cache_dir, index=False)
            except Exception as e:
                logger.error('df_chg save error', exc_info=e)
        else:
            try:
                self.cache.to_csv(self.cache_dir, index=False)
            except Exception as e:
                logger.error('df_chg save error', exc_info=e)

    @staticmethod
    def show_image(img_src, label):  # image2pixmap
        try:
            w = label.geometry().width()
            h = label.geometry().height()
            img_src_ = cv2.resize(img_src, (w, h), interpolation=cv2.INTER_AREA)
            frame = cv2.cvtColor(img_src_, cv2.COLOR_BGR2RGB)
            img = QImage(frame.data, frame.shape[1], frame.shape[0], frame.shape[2] * frame.shape[1],
                         QImage.Format_RGB888)
            label.setPixmap(QPixmap.fromImage(img))
        except Exception as e:
            logger.warning('show_image error', exc_info=e)

    def set_frame(self, e):  # lbl_frame double click event  프레임 더블클릭 시
        if not self.btn_pp.isChecked() and self.btn_pp.isEnabled() and self.det_thread.mode == 'video':
            dlg = QInputDialog(self)
            dlg.setWindowIcon(icon)
            dlg.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint)
            dlg.setInputMode(QInputDialog.IntInput)
            dlg.setWindowTitle('프레임 이동')
            dlg.setLabelText("원하는 프레임으로 이동합니다.")
            dlg.resize(500, 100)
            dlg.setIntRange(1, int(self.lbl_frame.text().split('/')[1]))
            dlg.setIntValue(int(self.lbl_frame.text().split('/')[0]))
            ok = dlg.exec_()
            num = dlg.intValue()
            if ok:
                try:
                    self.lbl_frame.setText(str(num) + '/' + self.lbl_frame.text().split('/')[1])
                    self.sliframe.setValue(num)
                except Exception as e:
                    logger.warning('set_frame error', exc_info=e)

    def send_frame(self, frame):  # show current lbl_frame  현재 프레임을 나타냅니다
        frames = self.lbl_frame.text().split('/')[1]
        try:
            self.sliframe.setValue(frame)
            self.lbl_frame.setText(str(frame) + '/' + str(frames))
        except Exception as e:
            logger.warning('send_frame error', exc_info=e)

    def send_frames(self, frames):  # show total frames  총 프레임을 나타냅니다
        try:
            self.sliframe.setRange(1, frames)
            self.lbl_frame.setText('1' + '/' + str(frames))
        except Exception as e:
            logger.warning('send_frames error', exc_info=e)

    def frame_chg(self):  # move slider to set lbl_frame  프레임 슬라이더를 움직여 동영상의 프레임을 맞춥니다
        num = int(self.sliframe.value())
        if num == self.sliframe.maximum():  # If slider set last lbl_frame of video
            self.btn_pp.setEnabled(False)  # play/pause btn set disable
        else:
            self.btn_pp.setEnabled(True)  # play/pause btn set enable
        if not self.det_thread.status:  # If video paused
            try:
                self.det_thread.frame = num  # set current lbl_frame to slider value
                if self.det_thread.vid_cap is not None:
                    self.det_thread.vid_cap.set(cv2.CAP_PROP_POS_FRAMES, num)  # set video to current lbl_frame
                self.lbl_frame.setText(str(num) + '/' + self.lbl_frame.text().split('/')[1])
            except Exception as e:
                logger.warning('frame_chg error', exc_info=e)

    def frame_reset(self):  # 프레임 초기화
        try:
            self.sliframe.setValue(0)
            self.sliframe.setRange(0, 0)
        except Exception as e:
            logger.warning('frame_reset error', exc_info=e)

    def pp(self):  # play/pause  재생/일시정지
        if self.btn_pp.isChecked():
            self.det_thread.play()
        else:
            self.det_thread.pause()

    def edit_time(self):
        self.hours = self.timebox.value()
        if self.det_thread.isRunning():
            adj_time = self.raw_time + timedelta(hours=self.hours)
            self.sample_table.setItem(0, 1, QTableWidgetItem(str(adj_time.year)))
            self.sample_table.setItem(0, 2, QTableWidgetItem(str(adj_time.month)))
            self.sample_table.setItem(0, 3, QTableWidgetItem(str(adj_time.day)))
            self.sample_table.setItem(0, 4, QTableWidgetItem(str(adj_time.hour)))
            self.sample_table.setItem(0, 5, QTableWidgetItem(str(adj_time.minute)))

    def filetime(self, timedata):  # Main tab / tableview datetime  메인탭의 테이블에 날짜와 시간을 나타냅니다
        self.raw_time = timedata
        adj_time = self.raw_time + timedelta(hours=self.hours)
        self.sample_table.setItem(0, 1, QTableWidgetItem(str(adj_time.year)))
        self.sample_table.setItem(0, 2, QTableWidgetItem(str(adj_time.month)))
        self.sample_table.setItem(0, 3, QTableWidgetItem(str(adj_time.day)))
        self.sample_table.setItem(0, 4, QTableWidgetItem(str(adj_time.hour)))
        self.sample_table.setItem(0, 5, QTableWidgetItem(str(adj_time.minute)))

    def add_column(self):  # add columns to table  칼럼 추가
        dlg = QInputDialog(self)
        dlg.setWindowIcon(icon)
        dlg.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint)
        dlg.setInputMode(QInputDialog.TextInput)
        dlg.setWindowTitle('칼럼 추가')
        dlg.setLabelText('추가할 칼럼이름을 입력하세요\n여러 개 입력 시 "/"키로 구분해 주세요')
        dlg.resize(700, 100)
        dlg.resize(700, 100)
        ok = dlg.exec_()
        text = [x for x in dlg.textValue().split('/') if x not in self.app_data['columns']]  # 중복칼럼제거
        if ok and text:
            self.app_data['columns'] += text
            self.table.setColumnCount(len(self.app_data['columns']))
            self.table.setHorizontalHeaderLabels(self.app_data['columns'])
            self.sample_table.setColumnCount(len(self.app_data['columns']))
            self.sample_table.setHorizontalHeaderLabels(self.app_data['columns'])
            try:
                with open(self.app_data_dir, 'w', encoding='utf-8') as f:
                    yaml.dump(self.app_data, f)
            except Exception as e:
                logger.error('add_column error', exc_info=e)

    def remove_column(self):  # remove columns to table  칼럼 제거
        dlg = QInputDialog(self)
        dlg.setWindowIcon(icon)
        dlg.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint)
        dlg.setInputMode(QInputDialog.TextInput)
        dlg.setWindowTitle('칼럼 제거')
        dlg.setLabelText('제거할 칼럼이름을 입력하세요\n여러 개 입력 시 "/"키로 구분해 주세요\n* 기본값은 제거할 수 없습니다.\n* Result 탭의 칼럼도 모두 사라집니다.')
        dlg.resize(700, 100)
        ok = dlg.exec_()
        text = dlg.textValue().split('/')
        if ok:
            reply = QMessageBox.question(self, '칼럼 제거',
                                         f"{', '.join(text)}칼럼 제거를 시도합니다.\n* Result 탭의 칼럼데이터도 모두 사라집니다.",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                success = []
                fail = []
                for name in text:
                    try:
                        self.app_data['columns'].remove(name)
                    except ValueError:
                        fail.append(name)
                    else:
                        success.append(name)
                self.table.setColumnCount(len(self.app_data['columns']))
                self.table.setHorizontalHeaderLabels(self.app_data['columns'])
                self.sample_table.setColumnCount(len(self.app_data['columns']))
                self.sample_table.setHorizontalHeaderLabels(self.app_data['columns'])
                try:
                    with open(self.app_data_dir, 'w', encoding='utf-8') as file:
                        yaml.dump(self.app_data, file)
                except Exception as e:
                    logger.error('remove_column error', exc_info=e)
                QMessageBox.information(self, '칼럼 제거 결과',
                                        f"칼럼 제거 성공 : {', '.join(success)}\n칼럼 제거 실패 : {', '.join(fail)}")

    # def save_txt(self):
    #     if self.chk_lbl.isChecked():
    #         self.lbl_txt.setEnabled(True)
    #         self.btn_txt.setEnabled(True)
    #         self.det_thread.save_txt = True
    #     else:
    #         self.lbl_txt.setEnabled(False)
    #         self.btn_txt.setEnabled(False)
    #         self.det_thread.save_txt = False

    def save_df(self):  # save result table to excel  결과 테이블을 엑셀로 저장합니다
        now = datetime.now().strftime('%Y%m%d%H%M')
        p = QFileDialog.getSaveFileName(self, 'Save File', now, ".xlsx(*.xlsx)")[0]
        col_count = self.table.columnCount()
        row_count = self.table.rowCount()
        headers = [str(self.table.horizontalHeaderItem(i).text()) for i in range(col_count)]
        df_list = []
        for row in range(row_count):
            df_list2 = []
            for col in range(col_count):
                table_item = '' if self.table.item(row, col) is None else self.table.item(row, col).text()
                df_list2.append('' if table_item == 'nan' else table_item)
            df_list.append(df_list2)
        df = pd.DataFrame(df_list, columns=headers)
        if p:
            try:
                df.to_excel(p, index=False)  # method 1
            except Exception as e:
                logger.error('save_df method 1 error', exc_info=e)
                try:
                    self.cache.to_excel(p, index=False)  # method 2
                except Exception as e:
                    logger.critical('save_df method 2 error', exc_info=e)
                    QMessageBox.critical(self, '저장 실패!', f'저장에 실패하였습니다! 잠시 후 다시 시도해주세요.\n에러코드:{type(e)}{e}',
                                         QMessageBox.Ok)
                else:
                    QMessageBox.warning(self, '불완전한 저장',
                                        f'에러가 발생하여 임시파일로 저장했습니다. 엑셀파일을 확인해주세요.\n해당 경고가 지속적으로 발생하면 문의바랍니다.\n{p}로 저장되었습니다.',
                                        QMessageBox.Ok)
                    self.save_status = True
            else:
                QMessageBox.information(self, '저장 완료', f'{p}로 저장되었습니다.', QMessageBox.Ok)
                self.save_status = True

    def rm_data(self):  # remove result table one line  입력한 숫자의 행을 지웁니다.
        row = self.table.rowCount()
        dlg = QInputDialog(self)
        dlg.setWindowIcon(icon)
        dlg.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint)
        dlg.setInputMode(QInputDialog.IntInput)
        dlg.setWindowTitle('행 삭제')
        dlg.setLabelText("입력한 숫자의 행을 지웁니다.")
        dlg.resize(500, 100)
        dlg.setIntRange(1, row)
        dlg.setIntValue(row)
        ok = dlg.exec_()
        num = dlg.intValue()
        if ok:
            try:
                self.table.removeRow(num - 1)
                self.cache = self.cache.drop(num - 1).reset_index(drop=True)
                self.cache.to_csv(self.cache_dir, index=False)
            except Exception as e:
                QMessageBox.warning(self, 'Warning!', f'에러코드{type(e)} {e}', QMessageBox.Ok)
                logger.error('rm_data error', exc_info=e)
            else:
                self.save_status = False

    def night(self, e):  # Main tab / auto judge noon or night  주/야간을 판단합니다.
        if e:
            self.night_btn.setChecked(True)
        else:
            self.noon_btn.setChecked(True)

    def animalcount(self, count):  # Main tab / view table for animal counts  개체 수를 나타냅니다
        self.sample_table.setItem(0, 9, QTableWidgetItem(str(count)))

    def path(self, p):  # Main tab / view table to path  파일명을 나타냅니다
        self.sample_table.setItem(0, 0, QTableWidgetItem(p.split('\\')[-1]))

    def status(self, e):  # status (1:video is running, 2:video is stopped, 3:image  상태 (1:비디오 실행 중, 2:비디오 멈춤, 3:이미지)
        logger.info(f'signal {e}')
        if e == 1:  # when video is running  비디오 실행 중
            self.btn_pp.setEnabled(True)
            self.btn_pp.setChecked(True)
            self.options.setEnabled(False)
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(False)
            self.sliframe.setEnabled(False)
            self.btn_submit.setEnabled(False)
            self.btn_add.setEnabled(False)
            self.btn_remove.setEnabled(False)
            self.btn_export.setEnabled(False)
            self.btn_add_li.setEnabled(False)
            self.btn_rm_file.setEnabled(False)
        elif e == 2:  # when video is stopped  비디오 멈춤
            self.btn_pp.setChecked(False)
            self.btn_prev.setEnabled(True)
            self.btn_next.setEnabled(True)
            self.sliframe.setEnabled(True)
            self.btn_submit.setEnabled(True)
            self.btn_add.setEnabled(True)
            self.btn_remove.setEnabled(True)
            self.btn_export.setEnabled(True)
            self.options.setEnabled(True)
            self.btn_add_li.setEnabled(True)
            self.btn_rm_file.setEnabled(True)
            if self.lbl_cnt.text().split('/')[0] == '1':  # If progress is first file
                self.btn_prev.setEnabled(False)
            if self.lbl_cnt.text().split('/')[0] == self.lbl_cnt.text().split('/')[1]:  # If progress is last file
                self.btn_next.setEnabled(False)
            if self.sliframe.value() == self.sliframe.maximum():
                self.btn_pp.setEnabled(False)
            else:
                self.btn_pp.setEnabled(True)
        elif e == 3:  # when image  이미지일 때
            self.btn_pp.setEnabled(False)
            self.btn_pp.setChecked(False)
            self.sliframe.setEnabled(False)
            self.btn_submit.setEnabled(True)
            self.btn_add.setEnabled(True)
            self.btn_remove.setEnabled(True)
            self.btn_export.setEnabled(True)
            self.options.setEnabled(True)
            self.btn_add_li.setEnabled(True)
            self.btn_rm_file.setEnabled(True)
            if self.lbl_cnt.text().split('/')[0] == '1':  # If progress is first file
                self.btn_prev.setEnabled(False)
            if self.lbl_cnt.text().split('/')[0] == self.lbl_cnt.text().split('/')[1]:  # If progress is last file
                self.btn_next.setEnabled(False)
            QTest.qWait(200)
            self.btn_prev.setEnabled(True)
            self.btn_next.setEnabled(True)


# class GroupBox(QGroupBox):  # reverse gbx
#     def paintEvent(self, event):
#         painter = QStylePainter(self)
#         option = QStyleOptionGroupBox()
#         self.initStyleOption(option)
#         if self.isCheckable():
#             option.state &= ~QStyle.State_Off & ~QStyle.State_On
#             option.state |= (
#                 QStyle.State_Off
#                 if self.isChecked()
#                 else QStyle.State_On
#             )
#         painter.drawComplexControl(QStyle.CC_GroupBox, option)

ex = MyApp()
sys.exit(app.exec_())

# if __name__ == '__main__':
# app = QApplication(sys.argv)
# screen = app.primaryScreen()
# size = screen.size()
# w = size.width()
# h = size.height()
# scal = (h - 200) / h
# ex = MyApp()
# ex.move(0, 0)
# ex.setFixedSize(int(scal * w), h - 200)
# ex.setGeometry(int(0.1 * w), int(0.1 * h), int(0.8 * w), int(0.8 * h))
# ex.setGeometry(50,50, 1000, 600)
# ex.showMaximized()
# ex.showFullScreen()
# ex2 = StartWidget()
# sys.exit(app.exec_())
