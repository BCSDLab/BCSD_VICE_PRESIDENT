import os
import sys
from dotenv import load_dotenv
from pyhwpx import Hwp
from hwp.image_packer import pack

load_dotenv()
DEBUG = os.getenv('DEBUG', 'True').lower() == 'true'

def run(data, t_path, o_path):
    hwp = Hwp(visible=DEBUG)
    try:
        hwp.register_module()
        
        if not hwp.open(t_path, arg='versionwarning:False'):
            raise FileNotFoundError(f"입력 파일 열기 실패: {t_path}")
        
        hwp.save_as(o_path)

        pre_have = -1
        for ctrl in hwp.ctrl_list:
            if ctrl.CtrlID != 'tbl': continue
            prop = ctrl.Properties
            prop.SetItem('TreatAsChar', True)
            ctrl.Properties = prop
            pre_have += 1
        
        hwp.MoveDocBegin()
        hwp.MovePageDown()

        hwp.get_into_nth_table(1)
        hwp.DeleteLineEnd()

        hwp.TableLowerCell()
        hwp.Delete()

        hwp.SelectCtrlFront()
        hwp.Copy()

        hwp.UnSelectCtrl()

        for _ in range(len(data) - pre_have):
            hwp.MoveDocEnd()
            hwp.Paste()
        
        hwp.MoveDocBegin()
        hwp.MovePageDown()

        i = 1
        for idx, row in data.iterrows():
            hwp.get_into_nth_table(i)
            i += 1

            hwp.insert_text(f'{idx + 1}. {row["종류"]}')
            hwp.TableLowerCell()
            if not row['img_paths']:
                print(f"{idx}. 증빙 자료 누락. 확인 필요")
                continue

            pack(hwp, row['img_paths'])

        hwp.save()

    except Exception as e:
        print(f"err: {e}", file=sys.stderr)
    finally:
        if hwp and not DEBUG: hwp.quit()
