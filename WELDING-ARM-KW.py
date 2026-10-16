# -*- coding: utf-8 -*-
"""
WELDING-ARM-KW : 용접 암 측정 데이터(DATA-KW.xlsx) -> 타겟 스펙을 만족하는 시료의 전장 / 일자-Y Min·Max 추출기
------------------------------------------------------------
- 타겟 항목(단축 위치도 0~1.0, 장축 위치도 0~1.0, 단축 윤곽도 0~0.8)을 '모두 동시에' 만족하는
  시료만 골라서, 변수 항목(전장, 일자-Y)의 Min / Max (측정값과 편차 둘 다)와 해당 시료
  (파일명, 엑셀 행)를 찾는다. 스펙 경계값은 포함한다.
- 타겟 항목 열은 1행(구분 행)의 '단축 위치도 / 장축 위치도 / 단축 윤곽도' 라벨 위치로,
  변수 항목 열은 2행 머리글('전장', '3-1 일자 Y')로 찾는다
- 엑셀 하단의 MIN / MAX / AVG 요약 행, '0.8 미만' 같은 구분 행은 시료가 아니므로
  파일명 패턴(기본: '.xls' 또는 'QMS' 포함)으로 걸러내서 계산에서 제외한다
- UI/로그인/임시비밀번호 인프라/사이드바 업로드/상단 탭 구조는 WELDING-ARM-KW 원본과
  동일하게 유지한다 (차콜+앰버 테마)
"""

import io
import json
import base64
from datetime import datetime, timedelta
from collections import Counter

import bisect
import math
import re as _re
import numpy as np
import altair as alt
import streamlit as st
import pandas as pd
from openpyxl.utils import get_column_letter

# ============================================================
# 0. 기본 설정 / 스타일 (BA-CAPA 콘솔과 동일한 테마 + 로그인 방식)
# ============================================================
st.set_page_config(page_title="WELDING-ARM-KW | 스펙 만족 시료 Min/Max 분석", layout="wide", initial_sidebar_state="expanded")

OWNER_PASSWORD = "nt1234"

# ILJIN Corp CI 로고 (PNG, base64 인라인 - 별도 파일 배포 없이 앱 하나로 동작)
ILJIN_LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAz4AAAC8CAYAAAC5W3uWAAAKMWlDQ1BJQ0MgUHJvZmlsZQAAeJydlndUU9kWh8+9N71QkhCKlNBraFICSA29SJEuKjEJEErAkAAiNkRUcERRkaYIMijggKNDkbEiioUBUbHrBBlE1HFwFBuWSWStGd+8ee/Nm98f935rn73P3Wfvfda6AJD8gwXCTFgJgAyhWBTh58WIjYtnYAcBDPAAA2wA4HCzs0IW+EYCmQJ82IxsmRP4F726DiD5+yrTP4zBAP+flLlZIjEAUJiM5/L42VwZF8k4PVecJbdPyZi2NE3OMErOIlmCMlaTc/IsW3z2mWUPOfMyhDwZy3PO4mXw5Nwn4405Er6MkWAZF+cI+LkyviZjg3RJhkDGb+SxGXxONgAoktwu5nNTZGwtY5IoMoIt43kA4EjJX/DSL1jMzxPLD8XOzFouEiSniBkmXFOGjZMTi+HPz03ni8XMMA43jSPiMdiZGVkc4XIAZs/8WRR5bRmyIjvYODk4MG0tbb4o1H9d/JuS93aWXoR/7hlEH/jD9ld+mQ0AsKZltdn6h21pFQBd6wFQu/2HzWAvAIqyvnUOfXEeunxeUsTiLGcrq9zcXEsBn2spL+jv+p8Of0NffM9Svt3v5WF485M4knQxQ143bmZ6pkTEyM7icPkM5p+H+B8H/nUeFhH8JL6IL5RFRMumTCBMlrVbyBOIBZlChkD4n5r4D8P+pNm5lona+BHQllgCpSEaQH4eACgqESAJe2Qr0O99C8ZHA/nNi9GZmJ37z4L+fVe4TP7IFiR/jmNHRDK4ElHO7Jr8WgI0IABFQAPqQBvoAxPABLbAEbgAD+ADAkEoiARxYDHgghSQAUQgFxSAtaAYlIKtYCeoBnWgETSDNnAYdIFj4DQ4By6By2AE3AFSMA6egCnwCsxAEISFyBAVUod0IEPIHLKFWJAb5AMFQxFQHJQIJUNCSAIVQOugUqgcqobqoWboW+godBq6AA1Dt6BRaBL6FXoHIzAJpsFasBFsBbNgTzgIjoQXwcnwMjgfLoK3wJVwA3wQ7oRPw5fgEVgKP4GnEYAQETqiizARFsJGQpF4JAkRIauQEqQCaUDakB6kH7mKSJGnyFsUBkVFMVBMlAvKHxWF4qKWoVahNqOqUQdQnag+1FXUKGoK9RFNRmuizdHO6AB0LDoZnYsuRlegm9Ad6LPoEfQ4+hUGg6FjjDGOGH9MHCYVswKzGbMb0445hRnGjGGmsVisOtYc64oNxXKwYmwxtgp7EHsSewU7jn2DI+J0cLY4X1w8TogrxFXgWnAncFdwE7gZvBLeEO+MD8Xz8MvxZfhGfA9+CD+OnyEoE4wJroRIQiphLaGS0EY4S7hLeEEkEvWITsRwooC4hlhJPEQ8TxwlviVRSGYkNimBJCFtIe0nnSLdIr0gk8lGZA9yPFlM3kJuJp8h3ye/UaAqWCoEKPAUVivUKHQqXFF4pohXNFT0VFysmK9YoXhEcUjxqRJeyUiJrcRRWqVUo3RU6YbStDJV2UY5VDlDebNyi/IF5UcULMWI4kPhUYoo+yhnKGNUhKpPZVO51HXURupZ6jgNQzOmBdBSaaW0b2iDtCkVioqdSrRKnkqNynEVKR2hG9ED6On0Mvph+nX6O1UtVU9Vvuom1TbVK6qv1eaoeajx1UrU2tVG1N6pM9R91NPUt6l3qd/TQGmYaYRr5Grs0Tir8XQObY7LHO6ckjmH59zWhDXNNCM0V2ju0xzQnNbS1vLTytKq0jqj9VSbru2hnaq9Q/uE9qQOVcdNR6CzQ+ekzmOGCsOTkc6oZPQxpnQ1df11Jbr1uoO6M3rGelF6hXrtevf0Cfos/ST9Hfq9+lMGOgYhBgUGrQa3DfGGLMMUw12G/YavjYyNYow2GHUZPTJWMw4wzjduNb5rQjZxN1lm0mByzRRjyjJNM91tetkMNrM3SzGrMRsyh80dzAXmu82HLdAWThZCiwaLG0wS05OZw2xljlrSLYMtCy27LJ9ZGVjFW22z6rf6aG1vnW7daH3HhmITaFNo02Pzq62ZLde2xvbaXPJc37mr53bPfW5nbse322N3055qH2K/wb7X/oODo4PIoc1h0tHAMdGx1vEGi8YKY21mnXdCO3k5rXY65vTW2cFZ7HzY+RcXpkuaS4vLo3nG8/jzGueNueq5clzrXaVuDLdEt71uUnddd457g/sDD30PnkeTx4SnqWeq50HPZ17WXiKvDq/XbGf2SvYpb8Tbz7vEe9CH4hPlU+1z31fPN9m31XfKz95vhd8pf7R/kP82/xsBWgHcgOaAqUDHwJWBfUGkoAVB1UEPgs2CRcE9IXBIYMj2kLvzDecL53eFgtCA0O2h98KMw5aFfR+OCQ8Lrwl/GGETURDRv4C6YMmClgWvIr0iyyLvRJlESaJ6oxWjE6Kbo1/HeMeUx0hjrWJXxl6K04gTxHXHY+Oj45vipxf6LNy5cDzBPqE44foi40V5iy4s1licvvj4EsUlnCVHEtGJMYktie85oZwGzvTSgKW1S6e4bO4u7hOeB28Hb5Lvyi/nTyS5JpUnPUp2Td6ePJninlKR8lTAFlQLnqf6p9alvk4LTduf9ik9Jr09A5eRmHFUSBGmCfsytTPzMoezzLOKs6TLnJftXDYlChI1ZUPZi7K7xTTZz9SAxESyXjKa45ZTk/MmNzr3SJ5ynjBvYLnZ8k3LJ/J9879egVrBXdFboFuwtmB0pefK+lXQqqWrelfrry5aPb7Gb82BtYS1aWt/KLQuLC98uS5mXU+RVtGaorH1futbixWKRcU3NrhsqNuI2ijYOLhp7qaqTR9LeCUXS61LK0rfb+ZuvviVzVeVX33akrRlsMyhbM9WzFbh1uvb3LcdKFcuzy8f2x6yvXMHY0fJjpc7l+y8UGFXUbeLsEuyS1oZXNldZVC1tep9dUr1SI1XTXutZu2m2te7ebuv7PHY01anVVda926vYO/Ner/6zgajhop9mH05+x42Rjf2f836urlJo6m06cN+4X7pgYgDfc2Ozc0tmi1lrXCrpHXyYMLBy994f9Pdxmyrb6e3lx4ChySHHn+b+O31w0GHe4+wjrR9Z/hdbQe1o6QT6lzeOdWV0iXtjusePhp4tLfHpafje8vv9x/TPVZzXOV42QnCiaITn07mn5w+lXXq6enk02O9S3rvnIk9c60vvG/wbNDZ8+d8z53p9+w/ed71/LELzheOXmRd7LrkcKlzwH6g4wf7HzoGHQY7hxyHui87Xe4Znjd84or7ldNXva+euxZw7dLI/JHh61HXb95IuCG9ybv56Fb6ree3c27P3FlzF3235J7SvYr7mvcbfjT9sV3qID0+6j068GDBgztj3LEnP2X/9H686CH5YcWEzkTzI9tHxyZ9Jy8/Xvh4/EnWk5mnxT8r/1z7zOTZd794/DIwFTs1/lz0/NOvm1+ov9j/0u5l73TY9P1XGa9mXpe8UX9z4C3rbf+7mHcTM7nvse8rP5h+6PkY9PHup4xPn34D94Tz+6TMXDkAABbXSURBVHja7d15lBxV3YfxZ0hIIAkkgmyRJbKDhO1FQAQCxAsioCwKsoPbq0jBUXDllZ0DCvpCLu68IOIbAV8QPchiIaKIiKAsSlAIkEBYA2YjIQEy8/5RjUCYDNO3q2e6q5/POXM4Z+hbde+vKt39naq6t6unpwdJkiRJqrJlLIEkSZIkg48kSZIkGXwkSZIkyeAjSZIkSQYfSZIkSTL4SJIkSZLBR5IkSZIMPpIkSZIMPpIkSZJk8JEkSZIkg48kSZIkGXwkSZIkyeAjSZIkSQYfSZIkSTL4SJIkSZLBR5IkSZLBR5IkSZIMPpIkSZJk8JEkSZIkg48kSZIkGXwkSZIkyeAjSZIkSQYfSZIkSTL4SJIkSTL4SJIkSZLBR5IkSZIMPpIkSZJk8JEkSZIkg48kSZIkGXwkSZIkyeAjSZIkSQYfSZIkSQYfSZIkSTL4SJIkSZLBR5IkSZIMPpIkSZJk8JEkSZIkg48kSZIk9cPQJX/RddxNZwPLV3jM/yKG0wd8r1m+LXBIxc+nnxPD7/xnVWFZvh8woeKjvJAYpvr+ULobiOGGQThnz/cfbseYQQznWQZJr+rp6ek7+ACfAUZXuAbTgdMHYb+bAsdX/PyaBhh8qm1CB5zH1wBTB3ifnfD+MBu4YRD2e7z/bDtIli8khgsthKTeeKubJEmqim+R5dtbBkkGH0mSVGXLAleR5atZCkkGH0mSVGVjgcvJ8iGWQpLBR5IkVdkuwDmWQZLBR5IkVd2JZPkBlkGSwUeSJFXdj8jyjS2DJIOPJEmqslHA1WT5KEshyeAjSZKqbBPgYrK8y1JIBh9JkqQq+wjwOcsgGXwkSZKq7htk+QTLIBl8JEmSqmwIcAVZPtZSSAYfSZKkKlsN+BlZPsxSSAYfSZKkKtsBOM8ySAYfSZKkqsvI8kMtg2TwkSRJqrofkOXjLYNk8JEkSaqyERSLm46xFJLBR5IkqcrWBy51cVPJ4CNJklR1HwS+ahkkg48kSVLVnU6WB8sgVdtQSyBJUlu5E5jcwv37AtBui4QuA/yULN+aGB7zFJMMPpIkafBNIYbzW7Z3WX5UGwYfgJWBq8jyHYlhkaeZVD3e6iZJklTYBoiWQTL4SJIkVd0nyfKPWQbJ4CNJklR13yHLt7YMksFHkiSpyoZTLG66kqWQDD6SJElVtg4wmSz3u5Jk8JEkSaq0PYBTLYNk8JEkSaq6r5Hle1kGyeAjSZJUdZ+0BJLBR5IkSZIMPpIkSZJk8JEkSZIkg48kSZIkGXwkSZIkyeAjSZIkyeAjSZIkSQYfSZIkSTL4SJIkSZLBR5IkSZIMPpIkSZJk8JEkSZIkg48kSZIkGXwkSZIkGXwkSZIkyeAjSZIkSQYfSZIkSWpJQy2BJEmS+pTlw4GtgPHAJsA4YCywEjAKGFF75TzgFWAm8AwwA5gKPAjcSwyPWUwZfCRJktRKYWdjYD9gT2BbYHg/Wo2u/XfcUrb5DHAbcBPwa2J42ELL4CNJkqSBDjujgEOBTwFbN2EPqwH7134gy6cBVwCXE8M9HgA1k8/4SJKq7BVLIPUz8GT5KcA04HtNCj29GQd8CbibLN/TA6Fm8oqPJKnK5vParTeS3hx4uoAjgXMorsYMpnkeEBl8JEmSVHboWRP4ETCxRXo0x4Mig48kSZLKDD17AJMpZmVrFbM9MGomn/GRJEnqrNCTAde1WOgBr/jI4CNJkqSSQs9pwKSW/A4Yw1wPkAw+kiRJajT0fBk4uUV7N9sDJIOPJEmSGg09RwFnt3APvc1NTefkBpIkSdUOPdsB32/Clh8CbgOmANOBZ4EXgeWAkcDKwPrAJsCOwBp9bGu2B0oGH0mSJKWGntHAlcCwkrb4MHAR8FNimF5nXzYE9gcOBjY3+MjgI0mSpLJEYO0StvMo8FXgSmLoTutJeJBiodRzalehTgAOoHj0wlvdZPCRJElSgix/H3B4CVs6FziVGBaUF8fCHcCBZPkmwCkUt8hJBh9JkiTVFXqGUkxb3Yh5wBHEcE3T+hnDA8BHyfKRHjQZfCRJklSvIykmFUg1F9i9dmWm+WKY7yFTszmdtSRJUpVk+TLAlxrYwsvA3gMWeiSDjyRJkhIcAGzQQPsTiOFWy6iq8VY3SZKkajmhgbbXE0OsdHWyfBiwKbAxsCawGrAsxZTfL1FMtDCj9vN34BFi6KnAuFcBdqa4BfLtwBDgceBXxHD/IPVpdWAbYCOKdZ6WA7oppjd/DrgHuJcYSpn1z+AjSZJUnS/1GwLbJbZ+ETimonXZjOJK2O7AtnV+B55Hlv8WuBH4GTHMbHJfvw28q5+v7iGGXfvYVhfwIeB4YJelvGoEcH8f29gA+GEdI1hMDBP72N4Y4OPAYcCW/dheN1l+C3AFcBkxJM8AaPCRJEmqjsMaaDuJGKZVKOwMrdXjWOA/GtjSCsAHaz8XkOXXAl8nhj81qedrABPqGOdwYljUy+83BS5uIAi/Fvzq6c+r4SaG2Uv8bghwHMX05aPr2NoywG61n9PI8jOB76asJ+UzPpIkSdVxaGK7hcD5FQo9+wD/AC5pMPQsaSiwL3A7WX5d7WpI2Z6p8/Wjehn/gcCdJYSelP4ArLpEf1YFbga+VWfoWdLqwIXAb8jytQw+kiRJnahYDHTdxNaTieHpCtRgDFl+OfBLYL0m721P4D6y/PO1W8rKUu/zLCssUYNPUNwWNqKU3hTPN82rs9XbX9efscAfKZ4vKssutfA53uAjSZLUeSY20PbHFQg9GwN3AAcN4F6XA74JTCbLR5S0zQV1vn7062rwAeD7TRjn3DpfP6zWnxFNDKHvAG6t3dJn8JEkSeog70ts9zjw+zYPPdsBtwMbDlIPPgrkZPmKgxAyRtZqsAZwWZO+379S5+tfvQo1iXJvNewt9F1Xu5XO4CNJktQhUm8lurGtp2subvH7FTBmkHuyA3B1bbrsgfTq/n4IrNQiR2UIWT6RYva2ZluHYhIHg48kSVLlFQ96vy2xdd7G414R+AWwcov0aCLFA/wDaQWyfHdgrybuo95gvBxw3gDWYC+y/AiDjyRJUvVt0UDbW9t43BcCjc6s9jDw61qAuoVi4cxGfLY2q1yqF+p8/RDg7CbXud4JF46lf2v0lOmct3rOyuAjSZLU/jZPbPccMTzVliPO8gnA4YmtXwLOBcYRw/rEsAcx7FtbDHRVitvWrmmgd5EsXz6xbb3P0+wDbN1iR+e9ffy/x4GTKJ79GUkMXcDytaB0IvBg4j7XoFgnaKlcwFSSJKn9jUtsd1+bhp6hQExs/QiwDzFM6T2yhB6KiRL2I8sPAi4Fhte5j3VqX+LPGIBqHNVGR+484GRieHGJmi8E7gXuJcsvAL4CnJ5yZpDl3ySGlw0+kiRJ1bRWYruH2nS8+wPjE9o9AexEDE/269UxXEGWv0AxJXO9d0qdSJafTwzzmlyL/vRrCsUtjdOB52tBblWKqyzPD1RcJYYL+1HzV4AzyPI5wAV17mNs7dy4wuAjSZJUTWMT281o0/F+LqFND3BQv0PPa1/Ef0WWf53iKkQ9VgSOppjSeTD0UKzPdC4x3D/Ix+t7/Qo9b6z7JLJ8V2DfOvd12NKCj8/4SJIktb81Ets93XYjzfJtgO0TWl5EDLcl7vV0imdT6nXsIFXpKWACMRzVAqHnKeCExLZfALrrbLM7WT7a4CNJklRNKyS2e64Nx/qRhDY9wFnJeyyeQUmZnnkDsnzzAa7PDGA7YmiV2frOJIYFiXWfClxXZ6thwPsNPpIkSdW0XGK7F9pwrPsltLmRGKY3uN/LgEUD1N9Ui4C9ieHxFjlWcykmh2jElQltdjX4SJIkVU36tMlQ/9TJgz3WDUlbt+fKhvcdwyzgNwkt9x7ACp1DDPe20BH7JTHMb3AbNya0mWDwkSRJqp7hDbSd12Zj3T6x3fUl7T9PaLMlWb7cANRmLmm34zXTL0oInM8Cj9bZaiOyfKTBR5IkqVoWN9B2VJuNdbuENlOJoaxJHO5MaDOUYtroZvsJMbTarYu/K2k79a431QVsYfCRJEmqksbWielqs9Fuk9Dm7hL3fw/1zzIG8O4BqM01LXasphHDzJK29WBCm40NPpIkSdXzcmK7MW02zvUS2kwpMWTOBx5JaLluk+vSDdzeYseqzGm0n0loM87gI0mSVD2ptzi1z61uWT4CWDmhZdmLtKbMmLZ2k6vzaAve5vZYidtKuVXR4CNJklRBzyS2W7ONxpgaHp4ouR9PD2Df+2t6Cx6vMoNPynpTqxp8JEmSqif1qsbYNhrj6ontnmyBWq/S5NrMrXjwSXlWaCWDjyRJUvWkfrl/ZxuNcURiuzkl92NBQpsVm1ybOS14vBaUuK2UCSUMPpLUh+eBf1oGSW0o9a/r49tojMMS2y0uuR8LE9os24HnZJlrRKWEqKEGH0nqXTdwMDE8ZSkktaH7EtutQ5aPbpMxpvaz7EVaU4LPqA48J8sMnC+VETYNPpJUOIUYcssgqU3d20Db7dpkjF0e5rayYJD37xUfSerFtcBZlkFSG5vawBfN0CZjTJ2ueflmf6Huh0UdeE6OGOT9v2jwkaQ3ehQ4ghh6LIWkthVDN/DnxNa7t8koUxdpHV5yP1JuW1vYgWflyEGuOQYfSXrjB9EBxDDLUkiqgJsS221Olm/SBuNLnblsuZL7kTLJwrwOPB/LnNAh5SrbbIOPJL3mGGK42zJIqohGnlM8og3Gl3qrW9mLtKasJ9SJwWe1Ere1QkKb5w0+klT4ITFcYhkkVchfSFvhHuBosnz5Fh9f6pTdq5fcj7ED2Pd2ts4g1/xN/xaG+h6hEr2DLN/SMlTaKhUZx11A5uGUVCkxLCbLrwSOSWi9GvAp4IIWHt+zZPki6n9m5x0tEHwe78AzclyJ20q5aveswUfNdGLtR2pl/wI+TAyLLIWkCrosMfgAfJEsv4QY5rbw+B4DNqizzUal7T3Lhybs/9V+d5p1Bzn4TF3yF97qJqV7nl6mSlRL6wYOIYbplkJSJcXwp96+8PXTWODMFh9hykKtm5e4/01Jmyzhbx14No6vBcUypEy+8YjBRyrHdOA9wG4UVxDUHk4jhhstg6SKO7+Btp8ly3dr4bHdldBmS7K8rKmVt05sd2cHnocjgK1K2lbKdh4w+EiN+zuwIzE8VPvL2s7AE5al5V0HnGEZJHWAi4GZiW2XAa4gy9du0bGlrFU0BNilpP2nbOdpYujU7wl7NLyFLN8YWLXOVrOIwVvdpAbdDuxMDDP+/ZsY7gd2AB60PC3rUeAwFymV1BFieJHGrvq8HbiZLF+zBUd3B2mLge5XwhfwZYEPJLS8pYPPxgNL2MZeCW3uXFqql9Q/1wHv63WxyxgeA3Yk7RK8mmsRxWQGLlIqqZNMAp5uoP16wC1k+fotFurmk7Ze0YFk+agG9/4B0mY3vaaDz8PxZPn2DW7jsLLCpsFH6p/LgH2JYUEfb8YzKZ75uclytZRjiOGvlkFSR4nhBeDLDW5lPeBOsny/pvY1y9cmy49scpBYATi2wZ5+MaHNS8ANHX42ntTAuTER2DKh5fUGHynN+cCRxPByPz5o5lFckv2ZZWsJ/0MMF1sGSR3qx8DNDW5jDHA1WX41Wb5uqb3L8tFk+anAFOq7nekqYH5ScMny1RL7uj/Fbe31uoYY5nT4ebg3Wb5XQs2HAN9I2N/jwL0GH6l+JxHD5+p6NiSGl4CDge9avkH1Fxr/654kta/is+tooIwv3vsBD5LlPyHLdyLL075DZvkytfbfAWYApwAjgdF1jGsOcEnC3t8GXFL7Ql1Pn9ds4DP9Ak9EqNX9nXW2OYO0WfR+srTvbS5gKvWuG/hPYrgo8cNmMXAMWT4TONlyDrhZFM/1LLQUkjo8/DxGlh8NXF3C1oYAh9Z+HifLbwZuo5g2+GGKmbQW1sLCcIqrRWsA6wCbUUxJvAuwci/bHpMQKD5T61M99gQuJss/0a87OYrQcz31zyoGcAcx/NGTECiejbqJLN+HGKa8Rc27aoH4K4n7+tHS/odXfKQ3e/Vh+Isa3lIMpwDHAc4mNnB6KBYpnWYpJAmI4efA6SVvdS3gSOAHwK3Ak8CLZHkPWd5DMfPa08DdFM/knAkcsJTQA/Vc8SnGNBX4dmLfjwD+QJZv08eX72XI8oMpJi3aLHE/J3jyvcG6wF1k+Wm93nKY5V1k+S4UExOckriPXxLDUmfZ9YqP9EZzgQ8Rwy0lfuDE2pWfHwPLWuKmO40YbrAMkvQGpwJrA0e1aP9GJ47pEIrpt+u1LcXEDX+ieA7qIWAesBKwBbA3xZWqVJcTw22edm+yPMWdMCeR5fdRXClcSHFFbUvSrqy93tl9/U+Dj/SamcAexHB36VuO4XKyfBbw89o/ejXH9bhIqST19jnUQ5Z/AhhWCwvtH3ximEWWfxr4vwb2u33tp0zPUNzt0enuApZ2VW0Ixa2PW5W4v8m1heWXylvdpMKjwA5NCT2vvUHfCOwKPG+5m2IaxSKl3ZZCknr9HFoMHE6xxk+rWb62QGi9Y7oK+GYLjeNl4KDaEhed7mTSFptNMRs48a1eZPCR4G/AjrX7hZv9oXMHsBPwhGUv1SLgAGL4l6WQpD4/h7qJ4Xjg07Uv6a1kdGK7LwGTW6D/3cBRxPA7TzQAnqKRNXzq8zFieMrgI/XtD8BOxPDkAH7oPEBxWf0flr80x7pIqSTV9Vn0/dpn0T/bPvgUV7KOAAZz3bZFwKHEMNmT6w3OB37d5H2cVZvA4y0ZfNTJrqV4pmfgFxaLYQawM/BnD0PDLillBj5J6rzw81eKB8rPrH1xb8/g82r4ieHjwPHA4gHu95PABGK43JPqTcelG/gw0Kw/Tk4ihv/q74sNPupUlwL7EsOCQXwzmAnsBuQejmR3A8dYBklK/ixaSAxfAzaqfTYuHsTejClhPJMoFr28Y4D6fBGwWe1WdvV+TOYBE4EybwHsBj5fu22z3ww+6kTfAo6uXRof7DeD+RRTZl7pYanbLIrnelykVJIa/zyaTgxHAetRTBYw0BPxPFF7Xy9jLPcBOwAHUzzH2wy/AN5NDJ8khlmeQG95TGbXws+pNH518W/Ae4nhv+tt2Nt01nMqXvq5g7Tflzqgtu3gLGI4t8XeDF4iyw8BnqWYbUdvrYdiBrdHKzKeTnh/GKyAWsW6Lmjx/s2rYN3nd1QAghPJ8q8A7wf2B/YA1mjC3h6mWEPnauA3xPByiePoBi4ny6+guLX8UGBfYJUGtvoAxR8q/5cYHvKjuO5jshg4jSy/lGJCisOBkXVs4T7gXIo1kl5J6UJXT48LykuS1Oq6urosggZHlncBG1JcRdkCeBcwjmJB1GFv0XouxexeTwCPAA9S/MX+bmJ4ZhDG8S7gPcD42pjGUixaOup1r5xFsbbfk8CUWn9/TwxPDFA/R1H/oqwvEMNzTezTPbVjX4+tiOGePrY5EtgdmFDb9lqvG/c8iquO91M8H3QtMdQ9EceSOcfgI0mSwUdK/UI8shYchlDcSfTK6764zm6J28rVmsFnACyZc4Z6JCVJan3+oVItaj593Qo4yfO2CrqOu6kS43ByA0mSJEmVZ/CRJEmSZPCRJEmSJIOPJEmSJBl8JEmSJMngI0mSJEkGH0mSJEky+EiSJEmSwUeSJEmSwUeSJEmSDD6SJEmSZPCRJEmSJIOPJEmSJBl8JEmSJMngI0mSJEkGH0mSJEky+EiSJEky+EiSJEmSwUeSJEmS2tFQSyBJkiSpD78FptXZZk6rDaKrp6fHQylJkiSp0rzVTZIkSZLBR5IkSZIMPpIkSZLU4v4fQf7UY4089dwAAAAASUVORK5CYII="

BG_DARK = "#000000"
SURFACE = "#000000"
SURFACE_ALT = "#000000"
BORDER = "#3a3a3a"
BORDER_STRONG = "#454545"
TEXT_MAIN = "#f2f2f2"
TEXT_MUTED = "#9a9a9a"
ACCENT_BLUE = "#f59e0b"
DANGER = "#ef4444"
WARNING = "#fb923c"
SUCCESS = "#10b981"

CUSTOM_CSS = f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&family=Inter:wght@400;500;600;700&display=swap');

    html {{ font-size: 90% !important; }}
    .stApp {{ background-color: {BG_DARK} !important; color: {TEXT_MAIN} !important; font-family: 'Inter', sans-serif; }}
    [data-testid="stAppViewContainer"] .main .block-container {{ max-width: 100% !important; padding-left: 2rem !important; padding-right: 2rem !important; }}

    h1, h2, h3, h4 {{ font-family: 'Inter', sans-serif; font-weight: 600 !important; color: {TEXT_MAIN} !important; }}
    div[data-testid="stMetricValue"] {{ color: {ACCENT_BLUE}; }}

    .block-title {{
        color: {ACCENT_BLUE}; font-size: 0.9rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.04em; margin-bottom: 12px; padding-bottom: 6px; border-bottom: 1px solid {BORDER_STRONG};
    }}
    .warn-banner {{
        background-color: {SURFACE}; border-left: 4px solid {WARNING}; padding: 10px 14px;
        border-radius: 4px; margin-bottom: 10px; font-size: 14px; color: {TEXT_MAIN};
    }}
    .kpi-card {{ background-color: {SURFACE}; border-radius: 6px; padding: 14px 16px; border: 1px solid {BORDER}; min-height: 118px; box-sizing: border-box; }}
    .kpi-label {{ font-size: 12px; color: {TEXT_MUTED}; margin: 0 0 4px 0; }}
    .kpi-value {{ font-size: 24px; font-weight: 600; margin: 0; color: {TEXT_MAIN}; }}

    .glass-card {{
        background: rgba(0,0,0,0.75); backdrop-filter: blur(6px);
        border: 1px solid {BORDER_STRONG}; border-radius: 12px;
    }}

    div.stButton > button, div.stDownloadButton > button {{
        height: 2.6rem !important; background: {SUCCESS} !important; color: #ffffff !important;
        border-radius: 4px !important; border: none !important; font-weight: 600; width: 100%;
    }}
    div.stButton > button:hover {{ opacity: 0.9; }}

    label, .stTextInput label, .stSelectbox label, .stFileUploader label, [data-testid="stWidgetLabel"] p {{ color: {TEXT_MUTED} !important; }}
    div[data-testid="stSelectbox"] > div > div, div[data-baseweb="select"] > div {{ background-color: #000000 !important; border-color: {BORDER_STRONG} !important; }}
    div[data-testid="stSelectbox"] span, div[data-baseweb="select"] span {{ color: {TEXT_MAIN} !important; }}
    [data-testid="stFileUploaderDropzone"] {{ background-color: {SURFACE} !important; border: 1px solid {BORDER} !important; }}
    [data-testid="stAlert"] p, [data-testid="stAlert"] * {{ color: {TEXT_MAIN} !important; }}
    .stCaption, [data-testid="stCaptionContainer"] {{ color: {TEXT_MUTED} !important; }}
    [data-testid="stExpander"] {{ border: 1.5px solid {ACCENT_BLUE} !important; border-radius: 8px !important; background: {SURFACE} !important; margin-top: 8px; margin-bottom: 8px; box-shadow: 0 0 10px rgba(245,158,11,0.10); }}
    [data-testid="stExpander"] summary {{ font-weight: 700 !important; color: {ACCENT_BLUE} !important; }}

    /* 사이드바 콘솔 스타일 (BA-CAPA와 동일) */
    section[data-testid="stSidebar"] {{ background-color: {SURFACE_ALT} !important; border-right: 1px solid {BORDER}; min-width: 300px !important; }}
    section[data-testid="stSidebar"] * {{ color: {TEXT_MAIN} !important; }}
    section[data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,0.2); }}
    section[data-testid="stSidebar"] div.stButton > button {{
        background-color: rgba(255,255,255,0.10); color: white; border: 1px solid rgba(255,255,255,0.3);
    }}

    /* 상단 탭 스타일 (BA-CAPA와 동일) */
    .stTabs [data-baseweb="tab-list"] {{ border-bottom: 2px solid {BORDER_STRONG}; gap: 5px; flex-wrap: wrap; }}
    .stTabs [data-baseweb="tab"], .stTabs button[data-baseweb="tab"], .stTabs [role="tab"] {{
        background-color: #000000 !important; border: 1px solid {BORDER_STRONG} !important;
        border-bottom: none !important; border-radius: 6px 6px 0 0 !important;
        color: {TEXT_MAIN} !important; font-weight: 700 !important; opacity: 1 !important;
        padding: 7px 12px !important; font-size: 0.82rem !important;
    }}
    .stTabs [data-baseweb="tab"] p {{ font-size: 0.82rem !important; }}
    .stTabs [data-baseweb="tab"] * {{ color: {TEXT_MAIN} !important; opacity: 1 !important; }}
    .stTabs [aria-selected="true"], .stTabs button[aria-selected="true"] {{
        background-color: #5c3d0a44 !important; border-color: {ACCENT_BLUE} !important; color: {ACCENT_BLUE} !important;
    }}
    .stTabs [aria-selected="true"] * {{ color: {ACCENT_BLUE} !important; }}
    [data-baseweb="tab"] {{ opacity: 1 !important; }}

    /* OCR 진행 상황 모달(팝업) 스타일 - 다크네이비 알약형 배지 */
    .progress-pill {{
        background: linear-gradient(180deg, #1e2a44, #111a2e);
        border: 1px solid #2d3f5c; border-radius: 999px;
        padding: 9px 18px; text-align: center; font-size: 0.85rem;
        color: #93c5fd; box-shadow: 0 4px 14px rgba(0,0,0,0.45);
        margin: 4px 0 4px 0;
    }}
    .progress-pill b {{ color: #60a5fa; }}
    .progress-pill .pct {{ color: #60a5fa; font-weight: 700; }}
    .progress-pill.done {{
        background: linear-gradient(180deg, #14351f, #0d2416);
        border: 1px solid #1f5c33; color: #86efac;
    }}
    .progress-pill.done b {{ color: {SUCCESS}; }}

    /* "처리 중" 박스는 st.dialog가 아니라 _run_with_progress()가 직접 그리는 div라서,
       그 스타일은 여기(전역 CSS)가 아니라 그 함수 안에서 인라인으로 정의한다 - 배경(바깥
       클릭해도 안 닫힘) 포함 전부 그 함수 책임. */
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ============================================================
# 0-1. 로그인 (관리자: OWNER_PASSWORD / 사용자: secrets.toml 의 user_password)
#   - 사용자 비밀번호는 secrets.toml 에 `user_password = "..."` 로 넣는다 (여러 개면 리스트).
#   - 관리자가 그 줄을 지우거나 바꾸면 이후 사용자 로그인은 막힌다 (관리자 로그인은 항상 가능).
#     이미 로그인해서 쓰고 있는 사용자는 브라우저를 새로고침/재접속할 때부터 막힌다.
#   - secrets.toml 자체가 없거나 user_password 가 비어 있으면 사용자 로그인은 전부 거부된다.
# ============================================================
import hmac


def _user_passwords() -> list:
    try:
        cfg = st.secrets.get("user_password", "")
    except Exception:
        return []
    if isinstance(cfg, str):
        cfg = [cfg]
    try:
        return [str(x).strip() for x in cfg if str(x).strip()]
    except TypeError:
        return []


def check_user_pwd(p: str) -> bool:
    p = (p or "").strip()
    if not p:
        return False
    return any(hmac.compare_digest(p.encode("utf-8"), c.encode("utf-8")) for c in _user_passwords())


def login_gate():
    if st.session_state.get("authenticated"):
        return True

    _, center, _ = st.columns([1, 1.8, 1])
    with center:
        st.markdown("<div style='height:60px;'></div>", unsafe_allow_html=True)
        st.markdown(
            f"""<div class='glass-card' style='padding:26px 32px; margin-top:12px;'>
                <div style='text-align:center; margin-bottom:14px;'>
                    <img src='data:image/png;base64,{ILJIN_LOGO_B64}' style='height:26px;'/>
                </div>
                <div style='color:{ACCENT_BLUE}; font-size:2.16rem; font-weight:700; letter-spacing:0.04em; text-align:center; margin-bottom:10px; line-height:1.25;'>WELDING ARM 측정 데이터 분석</div>
                <div style='text-align:center; margin:0 0 4px 0;'>
                    <span style='display:inline-block; color:{TEXT_MAIN}; font-size:1.35rem; font-weight:600; line-height:1.4; border-bottom:2px solid {ACCENT_BLUE}; padding-bottom:10px;'>타겟 스펙 만족 시료의 전장 · 일자-Y Min / Max 추출</span>
                </div>
                <div style='color:{TEXT_MUTED}; font-size:0.82rem; text-align:center; margin-top:12px;'>V1.0</div>
            </div>""",
            unsafe_allow_html=True,
        )
        st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
        pw_col, btn_col = st.columns([4, 1])
        with pw_col:
            pw = st.text_input("비밀번호", type="password", key="login_pw",
                                label_visibility="collapsed", placeholder="비밀번호 입력")
        with btn_col:
            clicked = st.button("시스템 접속", type="primary", use_container_width=True)
        if clicked:
            if pw == OWNER_PASSWORD:
                st.session_state["authenticated"] = True
                st.session_state["is_owner"] = True
                st.rerun()
            elif check_user_pwd(pw):
                st.session_state["authenticated"] = True
                st.session_state["is_owner"] = False
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")
    st.stop()
# ============================================================
# 1. 분석 설정값 / 데이터 처리 함수
# ============================================================

# 타겟 항목: 1행(구분 행)에 적힌 라벨 위치로 열을 찾는다 (2행 머리글은 '4 일자 119'처럼 구분이 안 됨).
DEFAULT_TARGET_RULES = {
    "단축 위치도": r"단축\s*위치도",
    "장축 위치도": r"장축\s*위치도",
    "단축 윤곽도": r"단축\s*윤곽도",
}
# 타겟 스펙 (하한, 상한) - 경계 포함. 사이드바에서 바꿀 수 있다.
DEFAULT_TARGET_SPECS = {
    "단축 위치도": (0.0, 1.0),
    "장축 위치도": (0.0, 1.0),
    "단축 윤곽도": (0.0, 0.8),
}
# 변수 항목: 2행 머리글 텍스트로 '측정값 열'을 찾고, 바로 오른쪽 열(머리글 비어 있음)을 '편차' 열로 본다.
# '일자-Y'는 '3-1 일자 Y 219.66' 열이다 ('3-2 휨 Y'와는 다른 항목).
DEFAULT_VARIABLE_RULES = {
    "전장": r"^\s*전장",
    "일자-Y": r"일자\s*Y",
}
# 시료 행으로 인정할 파일명 패턴 (MIN/MAX/AVG, '0.8 미만' 등 요약·구분 행 제외용)
DEFAULT_NAME_PATTERN = r"\.xls|QMS"

_SHORT_ID_RX = _re.compile(r"QMS_\d+")


def _short_id(name) -> str:
    m = _SHORT_ID_RX.search(str(name))
    return m.group(0) if m else str(name)


@st.cache_data(show_spinner=False)
def _list_sheets(file_bytes: bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


@st.cache_data(show_spinner=False)
def _load_raw_sheet(file_bytes: bytes, sheet_name):
    """머리글 위치를 코드에서 직접 찾으므로 header=None 으로 있는 그대로 읽는다."""
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)


def locate_columns(raw: pd.DataFrame, var_rules: dict, target_rules: dict):
    """'파일명' 머리글 행, 변수 항목(값 열/편차 열), 타겟 항목(값 열)의 위치를 찾는다."""
    header_row = None
    for i in range(min(len(raw), 15)):
        if str(raw.iat[i, 0]).strip() == "파일명":
            header_row = i
            break
    if header_row is None:
        raise ValueError("A열에서 '파일명' 머리글 행을 찾지 못했습니다. 업로드한 엑셀 형식을 확인해주세요.")

    header = raw.iloc[header_row]
    ncols = raw.shape[1]
    cols, missing = {}, []

    for key, pattern in var_rules.items():
        rx = _re.compile(pattern)
        found = None
        for c in range(1, ncols):
            v = header.iat[c]
            if isinstance(v, str) and rx.search(v.replace("\n", " ")):
                dev = c + 1 if (c + 1 < ncols and pd.isna(header.iat[c + 1])) else None
                found = {"value_col": c, "dev_col": dev, "header": v.strip(), "kind": "변수"}
                break
        if found is None:
            missing.append(key)
        else:
            cols[key] = found

    for key, pattern in target_rules.items():
        rx = _re.compile(pattern)
        found = None
        # 머리글 바로 위 행(구분 행)부터 위로, 마지막으로 머리글 행 자체도 확인한다.
        for r in list(range(header_row - 1, -1, -1)) + [header_row]:
            for c in range(1, ncols):
                v = raw.iat[r, c]
                if isinstance(v, str) and rx.search(v.replace("\n", " ")):
                    found = {"value_col": c, "dev_col": None, "header": v.strip(), "kind": "타겟"}
                    break
            if found:
                break
        if found is None:
            missing.append(key)
        else:
            cols[key] = found

    if missing:
        raise ValueError(f"머리글에서 다음 항목을 찾지 못했습니다: {', '.join(missing)} (분석 설정 관리 탭의 정규식을 확인해주세요)")
    return header_row, cols


# ---------- 양 끝단(파이프 기울기)을 포함한 전장 ----------
# 전장(= 휨 X − 일자 X)은 양쪽 끝단 '중심' 사이의 거리이다. 실제 끝단 파이프는 기울어져 있어서 원주 8점의 Y 값이 서로 다르므로,
# 지그를 만들 때는 중심 기준 전장에 기울어서 튀어나온 만큼을 더해야 제품이 들어간다.
#   왼쪽 끝단: Y-값_단1~8 (원주 8점의 Y 좌표, 약 -119)  -> 가장 왼쪽(최외측) = 가장 작은 Y, 안쪽(최내측) = 가장 큰 Y
#   오른쪽 끝단: Y-값_장1~8 (원주 8점의 Y 좌표, 약 156)  -> 가장 오른쪽(최외측) = 가장 큰 Y, 안쪽(최내측) = 가장 작은 Y
#   양 끝단 포함 전장 = 오른쪽 최외측 Y − 왼쪽 최외측 Y   (지그 전장 치수의 기준)
#   양 끝단 안쪽 전장 = 오른쪽 최내측 Y − 왼쪽 최내측 Y   (참고: 양 끝단이 가장 가까워지는 길이)
DEFAULT_END_RULES = {
    "왼쪽 끝단": r"Y\s*-?\s*값\s*_?\s*단\s*(\d)",
    "오른쪽 끝단": r"Y\s*-?\s*값\s*_?\s*장\s*(\d)",
}
END_ITEMS = ["전장", "양 끝단 포함 전장", "양 끝단 안쪽 전장", "기울기 증가량", "왼쪽 끝단 최외측 Y", "오른쪽 끝단 최외측 Y"]


def locate_end_columns(raw: pd.DataFrame, header_row: int, end_rules: dict) -> dict:
    """머리글에서 끝단 8점 열을 번호 순서대로 찾는다 (값 열 바로 옆의 편차 열은 머리글이 비어 있어 자동으로 제외된다)."""
    header = raw.iloc[header_row]
    out = {}
    for side, pattern in end_rules.items():
        rx = _re.compile(pattern)
        found = []
        for c in range(1, raw.shape[1]):
            v = header.iat[c]
            if isinstance(v, str):
                m = rx.search(v.replace("\n", " "))
                if m:
                    found.append((int(m.group(1)) if m.groups() and m.group(1) and m.group(1).isdigit() else 99, c))
        out[side] = [c for _, c in sorted(found)]
    return out


def _add_end_extents(out: pd.DataFrame, body: pd.DataFrame, endcols: dict):
    """시료 표(out)에 끝단 최외측/최내측 Y, 포함 전장, 안쪽 전장, 기울기 증가량 열을 붙이고, 점이 빠진 행(bool 배열)을 돌려준다."""
    import warnings
    left = np.column_stack([pd.to_numeric(body.iloc[:, c], errors="coerce").to_numpy(dtype=float) for c in endcols["왼쪽 끝단"]])
    right = np.column_stack([pd.to_numeric(body.iloc[:, c], errors="coerce").to_numpy(dtype=float) for c in endcols["오른쪽 끝단"]])
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", RuntimeWarning)  # 값이 전부 빈 행은 NaN 으로 처리
        out["왼쪽 끝단 최외측 Y"] = np.nanmin(left, axis=1)
        out["왼쪽 끝단 최내측 Y"] = np.nanmax(left, axis=1)
        out["오른쪽 끝단 최내측 Y"] = np.nanmin(right, axis=1)
        out["오른쪽 끝단 최외측 Y"] = np.nanmax(right, axis=1)
    out["양 끝단 포함 전장"] = out["오른쪽 끝단 최외측 Y"] - out["왼쪽 끝단 최외측 Y"]
    out["양 끝단 안쪽 전장"] = out["오른쪽 끝단 최내측 Y"] - out["왼쪽 끝단 최내측 Y"]
    out["기울기 증가량"] = out["양 끝단 포함 전장"] - out["전장"]
    return ~np.isfinite(left).all(axis=1) | ~np.isfinite(right).all(axis=1)


def end_extent_summary(sub: pd.DataFrame) -> pd.DataFrame:
    """합격(OK) 시료에서 끝단 포함 길이들의 Min/Max/평균과 해당 시료."""
    rows = []
    for name in END_ITEMS:
        if name not in sub.columns or sub[name].isna().all():
            continue
        s = sub[name]
        i_min, i_max = s.idxmin(), s.idxmax()
        rows.append({"항목": name, "Min": s.min(), "Min 시료": sub.at[i_min, "시료"], "Min 엑셀 행": int(sub.at[i_min, "엑셀 행"]),
                     "Max": s.max(), "Max 시료": sub.at[i_max, "시료"], "Max 엑셀 행": int(sub.at[i_max, "엑셀 행"]),
                     "평균": s.mean(), "Max-Min": s.max() - s.min()})
    return pd.DataFrame(rows)


def prepare_dataset(raw: pd.DataFrame, var_rules: dict, target_rules: dict, name_pattern: str, end_rules: dict = None):
    """시료 행만 골라 '엑셀 행 / 파일명 / 시료 / 전장 / 일자-Y (+편차) / 타겟 3종' 표로 만든다."""
    header_row, cols = locate_columns(raw, var_rules, target_rules)
    body = raw.iloc[header_row + 1:]

    out = pd.DataFrame({
        "엑셀 행": body.index + 1,  # raw 의 0부터 시작하는 index -> 실제 엑셀 행 번호
        "파일명": body.iloc[:, 0].astype(str).values,
    })
    out["시료"] = out["파일명"].map(_short_id)
    for key in list(var_rules) + list(target_rules):
        info = cols[key]
        out[key] = pd.to_numeric(body.iloc[:, info["value_col"]], errors="coerce").values
        if info["dev_col"] is not None:
            out[f"{key} 편차"] = pd.to_numeric(body.iloc[:, info["dev_col"]], errors="coerce").values

    endcols = locate_end_columns(raw, header_row, DEFAULT_END_RULES if end_rules is None else end_rules)
    has_end = bool(endcols) and all(len(v) > 0 for v in endcols.values())
    end_incomplete = _add_end_extents(out, body, endcols) if has_end else None

    key_cols = list(var_rules) + list(target_rules)
    is_sample = out["파일명"].str.contains(name_pattern, regex=True, na=False)
    has_number = out[key_cols].notna().any(axis=1)

    hdr_row = raw.iloc[header_row]
    info = {
        "header_row_excel": header_row + 1,
        "n_excluded_non_sample": int((~is_sample & has_number).sum()),  # MIN/MAX/AVG 등 숫자가 있는 비시료 행
        "n_incomplete": int((is_sample & out[key_cols].isna().any(axis=1)).sum()),  # 시료인데 값이 빠진 행
        "has_end": has_end,
        "end_cols": {side: {"cols": cs, "headers": [str(hdr_row.iat[c]).strip() for c in cs]} for side, cs in endcols.items()} if has_end else {},
        "n_end_incomplete": int((is_sample.to_numpy() & end_incomplete).sum()) if has_end else 0,
    }
    df = out[is_sample].dropna(subset=key_cols).reset_index(drop=True)
    info["n_samples"] = len(df)
    return df, cols, info


def analyze_spec(df: pd.DataFrame, specs: dict, var_keys):
    """타겟 항목 스펙(하한, 상한 - 경계 포함)을 '전부 동시에' 만족하는 시료만 골라서,
    변수 항목(var_keys)의 측정값/편차 Min/Max 와 해당 시료를 찾는다."""
    mask = pd.Series(True, index=df.index)
    status = []
    for key, (lo, hi) in specs.items():
        lo, hi = (lo, hi) if lo <= hi else (hi, lo)
        ok = (df[key] >= lo) & (df[key] <= hi)
        status.append({
            "타겟 항목": key, "스펙": f"{lo:g} ~ {hi:g}",
            "단독 합격 수": int(ok.sum()), "단독 합격률(%)": (ok.mean() * 100) if len(df) else 0.0,
            "전체 실측 Min": df[key].min() if len(df) else None, "전체 실측 Max": df[key].max() if len(df) else None,
        })
        mask &= ok
    subset = df[mask].sort_values(list(var_keys)[0]).reset_index(drop=True)
    status_df = pd.DataFrame(status)

    rows = []
    if not subset.empty:
        for key in var_keys:
            for label, col in (("측정값", key), ("편차", f"{key} 편차")):
                if col not in subset.columns:
                    continue
                i_min, i_max = subset[col].idxmin(), subset[col].idxmax()
                rows.append({
                    "항목": key, "구분": label,
                    "Min": subset.at[i_min, col], "Min 시료": subset.at[i_min, "시료"], "Min 엑셀 행": int(subset.at[i_min, "엑셀 행"]),
                    "Max": subset.at[i_max, col], "Max 시료": subset.at[i_max, "시료"], "Max 엑셀 행": int(subset.at[i_max, "엑셀 행"]),
                    "Max-Min": subset.at[i_max, col] - subset.at[i_min, col],
                })
    return pd.DataFrame(rows), subset, status_df


def build_result_xlsx(summary: pd.DataFrame, status_df: pd.DataFrame, subset: pd.DataFrame, params: dict, window=None, end_summary=None) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name="합격시료 Min_Max", index=False)
        if end_summary is not None and len(end_summary):
            end_summary.to_excel(w, sheet_name="양끝단 포함 전장 Min_Max", index=False)
        status_df.to_excel(w, sheet_name="타겟 합격 현황", index=False)
        if window is not None:
            rng, ver = window_frames(window)
            rng.to_excel(w, sheet_name="참고_합격 보장 범위", index=False)
            ver.to_excel(w, sheet_name="참고_합격 보장 범위", index=False, startrow=len(rng) + 3)
        subset.to_excel(w, sheet_name="합격 시료 데이터", index=False)
        pd.DataFrame(list(params.items()), columns=["조건", "값"]).to_excel(w, sheet_name="분석 조건", index=False)
        for ws in w.book.worksheets:
            for idx, col in enumerate(ws.iter_cols(min_row=1, max_row=1), start=1):
                header = str(col[0].value or "")
                ws.column_dimensions[get_column_letter(idx)].width = 60 if header == "파일명" else max(12, len(header) + 4)
    return bio.getvalue()


# ---------- 합격 보장 범위 (전장 × 일자-Y) ----------
def find_pass_window(df: pd.DataFrame, ok_flags, xkey: str, ykey: str):
    """두 변수(xkey, ykey)의 직사각형 범위 중, '불합격 시료가 하나도 들어오지 않으면서
    합격 시료를 가장 많이 포함하는' 범위를 정확하게 찾는다.

    -> 결과 범위(x 범위 AND y 범위)에 들어오는 시료는 (관측 데이터 기준) 전부 합격이다.
    알고리즘: 직사각형의 왼쪽 끝을 합격 시료의 x 값으로 정하고, 오른쪽으로 시료를 하나씩 넣으면서
    '불합격 y 값 사이의 빈 구간'에 들어간 합격 시료 수를 센다 (O(n^2 log n), 시료 천 개 수준은 1~2초)."""
    x = df[xkey].to_numpy(dtype=float)
    y = df[ykey].to_numpy(dtype=float)
    ok = np.asarray(ok_flags, dtype=bool)
    n_pass = int(ok.sum())
    if n_pass == 0:
        return None

    order = np.lexsort((ok, x))  # x 오름차순, 같은 x 에서는 불합격(False)을 먼저 넣어 경계에서 놓치지 않게 한다
    xs, ys, oks = x[order], y[order], ok[order]
    inf = float("inf")
    best = {"n": 0}
    for x1 in np.unique(xs[oks]):
        fy, py = [], []  # 이 x 띠 안의 불합격 y(정렬), 합격 y(정렬)
        for j in range(int(np.searchsorted(xs, x1, side="left")), len(xs)):
            yv = float(ys[j])
            if oks[j]:
                bisect.insort(py, yv)
                i = bisect.bisect_left(fy, yv)
                lo = fy[i - 1] if i > 0 else -inf
                hi = fy[i] if i < len(fy) else inf
                cnt = bisect.bisect_left(py, hi) - bisect.bisect_right(py, lo)  # (lo, hi) 사이 합격 시료 수
                if cnt > best["n"]:
                    best = {"n": cnt, "x1": float(x1), "x2": float(xs[j]), "lo": lo, "hi": hi}
            else:
                bisect.insort(fy, yv)
    if best["n"] == 0:
        return None

    inside = ok & (x >= best["x1"]) & (x <= best["x2"]) & (y > best["lo"]) & (y < best["hi"])
    xr = (float(x[inside].min()), float(x[inside].max()))
    yr = (float(y[inside].min()), float(y[inside].max()))
    box = (x >= xr[0]) & (x <= xr[1]) & (y >= yr[0]) & (y <= yr[1])  # 최종 범위 안의 '모든' 시료로 다시 검증
    return {"xkey": xkey, "ykey": ykey, "x": xr, "y": yr,
            "n_in": int(box.sum()), "n_pass_in": int((box & ok).sum()), "n_fail_in": int((box & ~ok).sum()),
            "n_pass": n_pass, "n_total": int(len(df))}


def window_frames(win: dict):
    """엑셀 저장용: (범위 표, 검증 표)."""
    rng = pd.DataFrame([
        {"항목": win["xkey"], "Min": win["x"][0], "Max": win["x"][1], "폭(Max-Min)": win["x"][1] - win["x"][0]},
        {"항목": win["ykey"], "Min": win["y"][0], "Max": win["y"][1], "폭(Max-Min)": win["y"][1] - win["y"][0]},
    ])
    ver = pd.DataFrame([{
        "두 범위에 동시에 들어오는 시료 수": win["n_in"], "그 중 합격": win["n_pass_in"], "그 중 불합격": win["n_fail_in"],
        "전체 합격 시료 수": win["n_pass"], "합격 시료 포함률(%)": win["n_pass_in"] / win["n_pass"] * 100 if win["n_pass"] else 0.0,
    }])
    return rng, ver


def window_tradeoff(df: pd.DataFrame, ok_flags, xkey: str, ykey: str,
                    rates=(0.0, 0.005, 0.01, 0.02, 0.03, 0.05), step: float = 0.01):
    """허용 불합격률(범위 안 시료 중 불합격 비율)별로 '합격 시료를 가장 많이 포함하는' 직사각형 범위를 찾는다.
    - 0% 는 정확 탐색(find_pass_window), 그 외는 step(mm) 격자 + 2차원 누적합으로 모든 직사각형을 벡터화해서 비교한다.
    - 격자로 찾은 뒤 실제 시료의 경계값으로 다시 좁히고 시료 기준으로 재검증하므로, 결과 불합격률은 허용률 이하다.
    - 허용률이 클수록 범위는 같거나 넓어지도록(합격 시료 포함 수가 줄지 않도록) 정리한다."""
    x = df[xkey].to_numpy(dtype=float)
    y = df[ykey].to_numpy(dtype=float)
    ok = np.asarray(ok_flags, dtype=bool)
    strict = find_pass_window(df, ok, xkey, ykey)
    if strict is None:
        return []

    k0x, k1x = int(np.floor(x.min() / step)), int(np.floor(x.max() / step)) + 1
    k0y, k1y = int(np.floor(y.min() / step)), int(np.floor(y.max() / step)) + 1
    xe, ye = np.arange(k0x, k1x + 1) * step, np.arange(k0y, k1y + 1) * step
    nx, ny = len(xe) - 1, len(ye) - 1
    cp = np.histogram2d(x[ok], y[ok], bins=[xe, ye])[0]
    cf = np.histogram2d(x[~ok], y[~ok], bins=[xe, ye])[0]
    sp, sf = np.zeros((nx + 1, ny + 1)), np.zeros((nx + 1, ny + 1))
    sp[1:, 1:] = cp.cumsum(0).cumsum(1)
    sf[1:, 1:] = cf.cumsum(0).cumsum(1)
    ai, bi = np.triu_indices(ny + 1, k=1)

    nonzero = [r for r in rates if r > 0]
    best = {r: (-np.inf, None) for r in nonzero}
    for i in range(nx):
        row_p, row_f = sp[i + 1:] - sp[i], sf[i + 1:] - sf[i]
        pc = row_p[:, bi] - row_p[:, ai]  # (오른쪽 끝 후보, y 구간 후보) 별 합격 시료 수
        fc = row_f[:, bi] - row_f[:, ai]  # 같은 위치의 불합격 시료 수
        score = pc * 10000 - fc           # 합격 시료 수 최대, 같으면 불합격이 적은 쪽
        for r in nonzero:
            feas = (fc * (1 - r) <= r * pc + 1e-9) & (pc > 0)
            if not feas.any():
                continue
            sc = np.where(feas, score, -np.inf)
            k = np.unravel_index(int(np.argmax(sc)), sc.shape)
            if sc[k] > best[r][0]:
                best[r] = (sc[k], (i, i + 1 + int(k[0]), int(ai[k[1]]), int(bi[k[1]])))

    def _from_grid(cell):
        i, j, a, b = cell
        inside = ok & (x >= xe[i]) & (x < xe[j]) & (y >= ye[a]) & (y < ye[b])
        if not inside.any():
            return None
        xr = (float(x[inside].min()), float(x[inside].max()))
        yr = (float(y[inside].min()), float(y[inside].max()))
        box = (x >= xr[0]) & (x <= xr[1]) & (y >= yr[0]) & (y <= yr[1])
        return {"xkey": xkey, "ykey": ykey, "x": xr, "y": yr, "n_in": int(box.sum()), "n_pass_in": int((box & ok).sum()),
                "n_fail_in": int((box & ~ok).sum()), "n_pass": int(ok.sum()), "n_total": int(len(df))}

    out = [dict(strict, rate=0.0)]
    for r in nonzero:
        cands = [dict(w) for w in out]  # 더 엄격한 허용률의 결과도 후보 (불합격률이 허용률 이하이므로 유효)
        if best[r][1] is not None:
            w = _from_grid(best[r][1])
            if w is not None:
                cands.append(w)
        valid = [c for c in cands if c["n_in"] and c["n_fail_in"] / c["n_in"] <= r + 1e-12]
        pick = max(valid, key=lambda c: (c["n_pass_in"], -c["n_fail_in"]))
        out.append(dict(pick, rate=r))
    return out


def gauge_design(win: dict, pct: float, wear_pct: float = 0.0, mm_override: dict = None) -> pd.DataFrame:
    """합격 범위 [L, U] 안쪽에 게이지 제작 공차 영역과 마모 여유를 두고 GO/NO-GO 설계 치수를 계산한다.
    - 하한측 설계 치수 = L + 마모 여유 + 제작 공차, 제작 공차 -제작공차/0 (실제 제작 치수는 L+마모여유 ~ 설계 치수)
    - 상한측 설계 치수 = U - 마모 여유 - 제작 공차, 제작 공차 0/+제작공차
    -> 게이지가 최악(공차 한계)으로 만들어져도 합격 범위 밖 부품은 통과시키지 않는다."""
    mm_override = mm_override or {}
    rows = []
    for key, (lo, hi) in ((win["xkey"], win["x"]), (win["ykey"], win["y"])):
        width = hi - lo
        tg = mm_override.get(key)
        tg = width * pct / 100 if tg is None else float(tg)
        wear = width * wear_pct / 100
        lo2, hi2 = lo + wear + tg, hi - wear - tg
        rows.append({"항목": key, "합격 범위 하한": lo, "합격 범위 상한": hi, "범위 폭": width,
                     "제작 공차(mm)": tg, "마모 여유(mm)": wear,
                     "하한측 설계 치수": lo2, "하한측 제작 공차": f"−{tg:.4f} / 0",
                     "상한측 설계 치수": hi2, "상한측 제작 공차": f"0 / +{tg:.4f}",
                     "게이지 통과 폭": hi2 - lo2, "설계 가능": bool(hi2 > lo2)})
    return pd.DataFrame(rows)


def gauge_check(df: pd.DataFrame, ok_flags, xkey: str, ykey: str, xr, yr) -> dict:
    """게이지(설계 치수 범위)로 실제 시료를 판정했을 때의 결과."""
    x, y = df[xkey].to_numpy(dtype=float), df[ykey].to_numpy(dtype=float)
    ok = np.asarray(ok_flags, dtype=bool)
    box = (x >= xr[0]) & (x <= xr[1]) & (y >= yr[0]) & (y <= yr[1])
    n_pass = int(ok.sum())
    return {"n_total": int(len(df)), "n_go": int(box.sum()), "n_go_pass": int((box & ok).sum()), "n_go_fail": int((box & ~ok).sum()),
            "n_pass": n_pass, "n_pass_lost": int((~box & ok).sum())}


def build_gauge_xlsx(trade_df: pd.DataFrame, design_df: pd.DataFrame, check: dict, params: dict) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        design_df.to_excel(w, sheet_name="게이지 설계 치수", index=False)
        pd.DataFrame([{
            "게이지 통과(GO) 시료 수": check["n_go"], "그 중 합격": check["n_go_pass"], "그 중 불합격(유출)": check["n_go_fail"],
            "게이지 통과율(%)": check["n_go"] / check["n_total"] * 100 if check["n_total"] else 0.0,
            "양품 중 게이지 NG 수": check["n_pass_lost"],
            "양품 손실률(%)": check["n_pass_lost"] / check["n_pass"] * 100 if check["n_pass"] else 0.0,
        }]).to_excel(w, sheet_name="게이지 판정 검증", index=False)
        trade_df.to_excel(w, sheet_name="허용 불합격률별 범위", index=False)
        pd.DataFrame(list(params.items()), columns=["조건", "값"]).to_excel(w, sheet_name="설계 조건", index=False)
        for ws in w.book.worksheets:
            for idx, col in enumerate(ws.iter_cols(min_row=1, max_row=1), start=1):
                ws.column_dimensions[get_column_letter(idx)].width = max(14, len(str(col[0].value or "")) + 4)
    return bio.getvalue()


def window_scatter_chart(df: pd.DataFrame, ok_flags, win: dict, inner: dict = None):
    d = pd.DataFrame({"x": df[win["xkey"]].values, "y": df[win["ykey"]].values, "sample": df["시료"].values,
                      "status": np.where(np.asarray(ok_flags, dtype=bool), "합격", "불합격")})
    pts = alt.Chart(d).mark_circle(size=30, opacity=0.6).encode(
        x=alt.X("x:Q", title=win["xkey"], scale=alt.Scale(zero=False)),
        y=alt.Y("y:Q", title=win["ykey"], scale=alt.Scale(zero=False)),
        color=alt.Color("status:N", title="합격 여부", scale=alt.Scale(domain=["합격", "불합격"], range=[SUCCESS, DANGER])),
        tooltip=[alt.Tooltip("sample:N", title="시료"), alt.Tooltip("x:Q", format=".4f", title=win["xkey"]),
                 alt.Tooltip("y:Q", format=".4f", title=win["ykey"]), alt.Tooltip("status:N", title="합격 여부")])
    layers = [pts, alt.Chart(pd.DataFrame({"x1": [win["x"][0]], "x2": [win["x"][1]], "y1": [win["y"][0]], "y2": [win["y"][1]]})).mark_rect(
        color="#ffffff", fillOpacity=0.10, stroke="#ffffff", strokeWidth=2).encode(x="x1:Q", x2="x2:Q", y="y1:Q", y2="y2:Q")]
    if inner is not None:  # 게이지 설계 범위(안쪽)를 점선 사각형으로 겹쳐 표시
        layers.append(alt.Chart(pd.DataFrame({"x1": [inner["x"][0]], "x2": [inner["x"][1]], "y1": [inner["y"][0]], "y2": [inner["y"][1]]})).mark_rect(
            color="#60a5fa", fillOpacity=0.15, stroke="#60a5fa", strokeWidth=2, strokeDash=[6, 4]).encode(x="x1:Q", x2="x2:Q", y="y1:Q", y2="y2:Q"))
    return _dark(alt.layer(*layers), 420)


# ============================================================
# 1-2. 제조업 품질 분석 함수 (Tab 5 품질 분석): 전장 구간별 합격률 / 공정능력 / 분포 / 상관 / 관리도
#   - 차트는 Streamlit 에 기본 포함된 Altair 로 그린다 (별도 패키지 설치 불필요).
#   - 차트 안의 필드 이름은 일부러 영문(ASCII)으로 쓴다 (Vega 표현식에서 한글 식별자 오류 방지).
# ============================================================

def _last_number(text):
    """머리글 텍스트의 마지막 숫자 (예: '전장 275.16' -> 275.16, '3-1 일자 Y 219.66' -> 219.66)."""
    nums = _re.findall(r"-?\d+(?:\.\d+)?", str(text))
    return float(nums[-1]) if nums else None


def pass_matrix(df: pd.DataFrame, specs: dict) -> pd.DataFrame:
    """타겟 항목별 합격 여부(경계 포함)와, 전부 만족한 '전체' 열."""
    out = pd.DataFrame(index=df.index)
    for key, (lo, hi) in specs.items():
        lo, hi = (lo, hi) if lo <= hi else (hi, lo)
        out[key] = (df[key] >= lo) & (df[key] <= hi)
    out["전체"] = out[list(specs)].all(axis=1)
    return out


# ---------- (4) 전장 구간별 합격률 ----------
def length_bin_table(df: pd.DataFrame, specs: dict, width: float) -> pd.DataFrame:
    flags = pass_matrix(df, specs)
    start = (np.floor(df["전장"].values / width + 1e-9) * width).round(6)
    t = pd.DataFrame({"start": start, "ok": flags["전체"].values})
    for k in specs:
        t[k] = flags[k].values
    g = t.groupby("start")
    out = pd.DataFrame({
        "시료 수": g.size(),
        "합격 수": g["ok"].sum().astype(int),
        "합격률(%)": g["ok"].mean() * 100,
    })
    for k in specs:
        out[f"{k} 단독 합격률(%)"] = g[k].mean() * 100
    out = out.reset_index().rename(columns={"start": "구간 시작"})
    out["구간 끝"] = (out["구간 시작"] + width).round(6)
    out["전장 구간"] = [f"{a:.2f} ~ {b:.2f}" for a, b in zip(out["구간 시작"], out["구간 끝"])]
    cols = ["전장 구간", "시료 수", "합격 수", "합격률(%)"] + [f"{k} 단독 합격률(%)" for k in specs] + ["구간 시작", "구간 끝"]
    return out[cols]


def length_threshold_curve(df: pd.DataFrame, specs: dict, step: float = 0.02) -> pd.DataFrame:
    """'전장 >= t' 인 시료만 남겼을 때의 시료 수와 합격률을 t 를 올려가며 계산한다."""
    flags = pass_matrix(df, specs)["전체"].values
    x = df["전장"].values
    grid = np.arange(np.floor(x.min() / step) * step, x.max(), step)
    rows = []
    for t in grid:
        m = x >= t - 1e-9
        n = int(m.sum())
        if n:
            rows.append({"전장 하한": round(float(t), 4), "시료 수": n, "합격률(%)": float(flags[m].mean() * 100)})
    return pd.DataFrame(rows)


def recommend_threshold(curve: pd.DataFrame, target_pct: float, min_n: int):
    """목표 합격률 이상이면서 시료 수가 min_n 이상인 가장 낮은 전장 하한. 없으면 None."""
    sub = curve[(curve["합격률(%)"] >= target_pct) & (curve["시료 수"] >= min_n)]
    return None if sub.empty else sub.iloc[0]


# ---------- (2) 공정능력 ----------
def _norm_sf(z: float) -> float:
    return 0.5 * math.erfc(z / math.sqrt(2))


def capability(x, lsl=None, usl=None) -> dict:
    """Cp / Cpk (σ = 표본 표준편차), 백분위 방식 Cpk, 정규분포 가정 예상 불량 PPM, 실측 불량 수.
    lsl 또는 usl 하나만 주면 한쪽 규격(Cpl 또는 Cpu 만)으로 계산한다."""
    x = pd.Series(x).dropna().astype(float)
    n = len(x)
    out = {"n": n, "mean": float(x.mean()) if n else np.nan, "std": float(x.std(ddof=1)) if n > 1 else np.nan}
    if n < 3 or not out["std"] or (lsl is None and usl is None):
        out.update(cp=np.nan, cpk=np.nan, cpk_pct=np.nan, ppm=np.nan, n_out=np.nan)
        return out
    m, sd = out["mean"], out["std"]
    cpl = (m - lsl) / (3 * sd) if lsl is not None else None
    cpu = (usl - m) / (3 * sd) if usl is not None else None
    out["cp"] = (usl - lsl) / (6 * sd) if (lsl is not None and usl is not None) else np.nan
    out["cpk"] = min(v for v in (cpl, cpu) if v is not None)

    p50, p_hi, p_lo = x.median(), x.quantile(0.99865), x.quantile(0.00135)
    pct = []
    if usl is not None and p_hi > p50:
        pct.append((usl - p50) / (p_hi - p50))
    if lsl is not None and p50 > p_lo:
        pct.append((p50 - lsl) / (p50 - p_lo))
    out["cpk_pct"] = min(pct) if pct else np.nan

    ppm = 0.0
    if lsl is not None:
        ppm += _norm_sf((m - lsl) / sd)
    if usl is not None:
        ppm += _norm_sf((usl - m) / sd)
    out["ppm"] = ppm * 1e6
    bad = pd.Series(False, index=x.index)
    if lsl is not None:
        bad |= x < lsl
    if usl is not None:
        bad |= x > usl
    out["n_out"] = int(bad.sum())
    return out


def cap_rating(c) -> str:
    if c is None or (isinstance(c, float) and np.isnan(c)):
        return "-"
    if c >= 1.67:
        return "우수 (≥1.67)"
    if c >= 1.33:
        return "양호 (≥1.33)"
    if c >= 1.00:
        return "보통 · 개선 필요 (≥1.00)"
    return "부족 (<1.00)"


# ---------- (1) 분포 ----------
def describe_item(x) -> dict:
    x = pd.Series(x).dropna().astype(float)
    n = len(x)
    if n < 3:
        return {"n": n}
    m, sd = x.mean(), x.std(ddof=1)
    d = x - m
    m2 = (d ** 2).mean()
    skew = (d ** 3).mean() / m2 ** 1.5 if m2 > 0 else 0.0
    kurt = (d ** 4).mean() / m2 ** 2 - 3 if m2 > 0 else 0.0
    jb = n / 6 * (skew ** 2 + kurt ** 2 / 4)  # Jarque-Bera (자유도 2 -> p = exp(-JB/2))
    return {"n": n, "평균": m, "표준편차": sd, "최소": x.min(), "중앙값": x.median(), "최대": x.max(),
            "왜도": skew, "첨도(초과)": kurt, "정규성 p값(JB)": math.exp(-jb / 2)}


_LINE_KINDS = ["LSL", "USL", "기준값", "CL", "UCL", "LCL", "목표", "Min", "Max"]
_LINE_COLORS = ["#fb7185", "#fb7185", "#10b981", "#10b981", "#a78bfa", "#a78bfa", "#10b981", "#60a5fa", "#fb7185"]  # 규격=분홍, 관리한계=보라 (막대 색과 구분)


def _line_color_scale(kinds=None):
    pairs = [(k, c) for k, c in zip(_LINE_KINDS, _LINE_COLORS) if kinds is None or k in kinds]
    return alt.Scale(domain=[k for k, _ in pairs], range=[c for _, c in pairs])


def _dark(chart, height=320):
    """검은 배경 테마에 맞춰 Altair 차트 색을 지정한다 (composite 차트의 최상위에서 한 번만 호출)."""
    return (chart.properties(height=height, background="#000000")
            .configure_axis(labelColor=TEXT_MUTED, titleColor=TEXT_MAIN, gridColor="#1f1f1f",
                            domainColor=BORDER_STRONG, tickColor=BORDER_STRONG)
            .configure_legend(labelColor=TEXT_MAIN, titleColor=TEXT_MAIN)
            .configure_title(color=TEXT_MAIN)
            .configure_view(strokeWidth=0))


def hist_chart(x, label, spec_lines=None):
    x = np.asarray(pd.Series(x).dropna(), dtype=float)
    nb = int(min(40, max(10, round(math.sqrt(len(x))))))
    counts, edges = np.histogram(x, bins=nb)
    bars = pd.DataFrame({"start": edges[:-1], "end": edges[1:], "count": counts, "zero": 0})
    m, sd = x.mean(), x.std(ddof=1)
    bw = edges[1] - edges[0]
    grid = np.linspace(edges[0], edges[-1], 200)
    curve = pd.DataFrame({"x": grid, "y": np.exp(-0.5 * ((grid - m) / sd) ** 2) / (sd * math.sqrt(2 * math.pi)) * len(x) * bw})
    bar = alt.Chart(bars).mark_rect(color=ACCENT_BLUE, opacity=0.85, stroke="#000000", strokeWidth=1).encode(
        x=alt.X("start:Q", title=label, scale=alt.Scale(zero=False), axis=alt.Axis(tickCount=10, labelOverlap=True)), x2="end:Q",
        y=alt.Y("count:Q", title="시료 수"), y2="zero:Q",
        tooltip=[alt.Tooltip("start:Q", format=".3f", title="구간 시작"), alt.Tooltip("end:Q", format=".3f", title="구간 끝"),
                 alt.Tooltip("count:Q", title="시료 수")])
    line = alt.Chart(curve).mark_line(color="#60a5fa", strokeWidth=2).encode(x="x:Q", y="y:Q")
    layers = [bar, line]
    if spec_lines:
        rules = pd.DataFrame([{"x": v, "kind": k} for k, v in spec_lines])
        layers.append(alt.Chart(rules).mark_rule(strokeDash=[6, 4], strokeWidth=2).encode(
            x="x:Q", color=alt.Color("kind:N", scale=_line_color_scale([k for k, _ in spec_lines]), legend=alt.Legend(title="기준선"))))
    return _dark(alt.layer(*layers), 320)


def box_compare_chart(all_vals, pass_vals, label):
    long = pd.concat([
        pd.DataFrame({"group": "전체 유효 시료", "v": pd.Series(all_vals).dropna().astype(float)}),
        pd.DataFrame({"group": "스펙 합격 시료", "v": pd.Series(pass_vals).dropna().astype(float)}),
    ])
    ch = alt.Chart(long).mark_boxplot(extent=1.5, size=26).encode(
        y=alt.Y("group:N", title=None),
        x=alt.X("v:Q", title=label, scale=alt.Scale(zero=False), axis=alt.Axis(tickCount=10, labelOverlap=True)),
        color=alt.Color("group:N", legend=None, scale=alt.Scale(domain=["전체 유효 시료", "스펙 합격 시료"], range=[ACCENT_BLUE, SUCCESS])))
    return _dark(ch, 170)


# ---------- (3) 상관 / 산점도 ----------
def corr_heatmap(df, items):
    c = df[items].corr()
    long = c.reset_index().melt(id_vars="index", var_name="y", value_name="r").rename(columns={"index": "x"})
    base = alt.Chart(long).encode(x=alt.X("x:N", title=None, sort=items), y=alt.Y("y:N", title=None, sort=items))
    rect = base.mark_rect().encode(color=alt.Color(
        "r:Q", scale=alt.Scale(domain=[-1, 0, 1], range=["#3b82f6", "#111111", "#f59e0b"]), legend=alt.Legend(title="r")))
    text = base.mark_text(fontSize=13).encode(
        text=alt.Text("r:Q", format=".2f"),
        color=alt.condition("abs(datum.r) > 0.5", alt.value("#ffffff"), alt.value("#9ca3af")))
    return _dark(alt.layer(rect, text), 330)


def scatter_chart(sdf, xlabel, ylabel, slope, intercept):
    xs = np.array([sdf["x"].min(), sdf["x"].max()])
    fit = pd.DataFrame({"x": xs, "y": slope * xs + intercept})
    pts = alt.Chart(sdf).mark_circle(size=34, opacity=0.55).encode(
        x=alt.X("x:Q", title=xlabel, scale=alt.Scale(zero=False)),
        y=alt.Y("y:Q", title=ylabel, scale=alt.Scale(zero=False)),
        color=alt.Color("status:N", title="합격 여부", scale=alt.Scale(domain=["합격", "불합격"], range=[SUCCESS, DANGER])),
        tooltip=[alt.Tooltip("sample:N", title="시료"), alt.Tooltip("x:Q", format=".4f", title=xlabel),
                 alt.Tooltip("y:Q", format=".4f", title=ylabel), alt.Tooltip("status:N", title="합격 여부")])
    line = alt.Chart(fit).mark_line(color="#ffffff", strokeDash=[6, 4], strokeWidth=2).encode(x="x:Q", y="y:Q")
    return _dark(alt.layer(pts, line), 380)


# ---------- (5) 런차트 / 관리도 (I-MR) ----------
def imr_stats(values) -> dict:
    v = np.asarray(values, dtype=float)
    n = len(v)
    mean = float(v.mean())
    mr = np.abs(np.diff(v))
    mrbar = float(mr.mean()) if len(mr) else 0.0
    ucl, lcl, mr_ucl = mean + 2.66 * mrbar, mean - 2.66 * mrbar, 3.267 * mrbar
    rule1 = (v > ucl) | (v < lcl)  # 관리한계 밖
    sign = np.sign(v - mean)
    rule2 = np.zeros(n, dtype=bool)  # 중심선 한쪽에 8점 이상 연속
    i = 0
    while i < n:
        j = i
        while j + 1 < n and sign[j + 1] == sign[i] and sign[i] != 0:
            j += 1
        if sign[i] != 0 and j - i + 1 >= 8:
            rule2[i:j + 1] = True
        i = j + 1
    return {"mean": mean, "mrbar": mrbar, "ucl": ucl, "lcl": lcl, "mr_ucl": mr_ucl,
            "rule1": rule1, "rule2": rule2, "mr": mr}


def imr_charts(order, values, label, st_info, spec_lines=None, window=30):
    v = np.asarray(values, dtype=float)
    d = pd.DataFrame({"row": np.asarray(order, dtype=float), "val": v,
                      "ooc": st_info["rule1"] | st_info["rule2"],
                      "ma": pd.Series(v).rolling(window, min_periods=max(3, window // 3)).mean().values})
    base = alt.Chart(d).encode(x=alt.X("row:Q", title="엑셀 행 (파일 순서)", scale=alt.Scale(zero=False)))
    line = base.mark_line(color="#6b7280", strokeWidth=1).encode(y=alt.Y("val:Q", title=label, scale=alt.Scale(zero=False)))
    pts = base.mark_point(filled=True, size=26).encode(
        y="val:Q", color=alt.condition("datum.ooc", alt.value(DANGER), alt.value(ACCENT_BLUE)),
        tooltip=[alt.Tooltip("row:Q", title="엑셀 행"), alt.Tooltip("val:Q", format=".4f", title=label)])
    ma = base.mark_line(color="#60a5fa", strokeWidth=2.5).encode(y="ma:Q")
    lines = [("CL", st_info["mean"]), ("UCL", st_info["ucl"]), ("LCL", st_info["lcl"])] + list(spec_lines or [])
    rules = alt.Chart(pd.DataFrame([{"y": y, "kind": k} for k, y in lines])).mark_rule(strokeDash=[6, 4], strokeWidth=1.8).encode(
        y="y:Q", color=alt.Color("kind:N", scale=_line_color_scale([k for k, _ in lines]), legend=alt.Legend(title="기준선")))
    i_chart = _dark(alt.layer(line, pts, ma, rules), 340)

    mrd = pd.DataFrame({"row": d["row"].values[1:], "mr": st_info["mr"]})
    mbase = alt.Chart(mrd).encode(x=alt.X("row:Q", title="엑셀 행 (파일 순서)", scale=alt.Scale(zero=False)))
    mline = mbase.mark_line(color="#6b7280", strokeWidth=1).encode(y=alt.Y("mr:Q", title="이동범위 |x(i)-x(i-1)|"))
    mpts = mbase.mark_point(filled=True, size=20, color=ACCENT_BLUE).encode(y="mr:Q")
    mrule = alt.Chart(pd.DataFrame([{"y": st_info["mr_ucl"], "kind": "UCL"}])).mark_rule(strokeDash=[6, 4], strokeWidth=1.8).encode(
        y="y:Q", color=alt.Color("kind:N", scale=_line_color_scale(), legend=None))
    mr_chart = _dark(alt.layer(mline, mpts, mrule), 220)
    return i_chart, mr_chart


# ============================================================
# 1-3. 조건 찾기: 입력(전장, 일자-Y) -> 결과(위치도 2종, 윤곽도) 예측 모델, 예측 정확도, 추천 조건, NG 시료 조정
#   - 모델: 입력 2개로 결과 3개를 각각 회귀 (선형 또는 2차: 곡률 + 상호작용 포함)
#   - OK 확률: 모델 예측값 + 실제 잔차(오차) 벡터를 재표본해서 '3개 스펙을 동시에' 만족할 확률을 구한다
#     (결과 3개가 서로 상관돼 있어도 반영되고, 정규분포 가정이 필요 없다)
#   - 정확도는 모두 교차검증(학습에 쓰지 않은 시료로 평가) 값이다
# ============================================================

alt.data_transformers.disable_max_rows()  # 히트맵 칸 수가 많아도 그릴 수 있게 (Altair 기본 5000행 제한 해제)


def _feat_matrix(x, y, norm, quad):
    mx, my, sx, sy = norm
    u = (np.asarray(x, dtype=float) - mx) / sx
    v = (np.asarray(y, dtype=float) - my) / sy
    cols = [np.ones_like(u), u, v]
    if quad:
        cols += [u * u, v * v, u * v]
    return np.column_stack(cols)


def fit_condition_model(x, y, tm, quad=True):
    x, y, tm = np.asarray(x, float), np.asarray(y, float), np.asarray(tm, float)
    norm = (float(x.mean()), float(y.mean()), float(x.std()) or 1.0, float(y.std()) or 1.0)
    X = _feat_matrix(x, y, norm, quad)
    beta, *_ = np.linalg.lstsq(X, tm, rcond=None)
    return {"quad": bool(quad), "norm": norm, "beta": beta, "resid": tm - X @ beta}


def predict_condition(model, x, y):
    return _feat_matrix(x, y, model["norm"], model["quad"]) @ model["beta"]


def pok_at(model, x, y, lo, hi):
    """(x, y) 조건에서 결과 3개가 동시에 스펙 안에 들어올 확률 (잔차 재표본)."""
    lo, hi = np.asarray(lo, float), np.asarray(hi, float)
    resid = model["resid"]
    pred = predict_condition(model, np.ravel(x), np.ravel(y))
    out = np.empty(len(pred))
    for s in range(0, len(pred), 400):
        blk = pred[s:s + 400]
        ok = np.ones((len(blk), len(resid)), dtype=bool)
        for t in range(blk.shape[1]):
            v = blk[:, t][:, None] + resid[None, :, t]
            ok &= (v >= lo[t]) & (v <= hi[t])
        out[s:s + 400] = ok.mean(axis=1)
    return out


@st.cache_data(show_spinner=False)
def cv_condition_model(x, y, tm, lo, hi, quad=True, k=5, seeds=(0, 1, 2)):
    """k-겹 교차검증 예측 정확도 + OK/NG 분류 정확도 + 확률 보정 표."""
    x, y, tm = np.asarray(x, float), np.asarray(y, float), np.asarray(tm, float)
    lo, hi = np.asarray(lo, float), np.asarray(hi, float)
    n = len(x)
    rmse = np.zeros(tm.shape[1])
    pred0 = pok0 = None
    for si, seed in enumerate(seeds):
        idx = np.random.default_rng(seed).permutation(n)
        pred, pok = np.zeros_like(tm), np.zeros(n)
        for f in np.array_split(idx, k):
            tr = np.setdiff1d(idx, f)
            m = fit_condition_model(x[tr], y[tr], tm[tr], quad)
            pred[f] = predict_condition(m, x[f], y[f])
            if si == 0:
                pok[f] = pok_at(m, x[f], y[f], lo, hi)
        rmse += np.sqrt(((tm - pred) ** 2).mean(axis=0))
        if si == 0:
            pred0, pok0 = pred, pok
    rmse /= len(seeds)
    sst = ((tm - tm.mean(axis=0)) ** 2).sum(axis=0)
    r2 = 1 - ((tm - pred0) ** 2).sum(axis=0) / sst

    ok_true = np.all((tm >= lo) & (tm <= hi), axis=1)
    pred_ok = pok0 >= 0.5
    tp, tn = int((pred_ok & ok_true).sum()), int((~pred_ok & ~ok_true).sum())
    fp, fn = int((pred_ok & ~ok_true).sum()), int((~pred_ok & ok_true).sum())
    n1, n0 = int(ok_true.sum()), int((~ok_true).sum())
    ranks = pd.Series(pok0).rank().to_numpy()
    auc = float((ranks[ok_true].sum() - n1 * (n1 + 1) / 2) / (n1 * n0)) if n1 and n0 else float("nan")
    edges = [0.0, 0.2, 0.4, 0.6, 0.8, 1.000001]
    cal = []
    for a, b in zip(edges[:-1], edges[1:]):
        m = (pok0 >= a) & (pok0 < b)
        if m.any():
            cal.append({"예측 OK 확률 구간": f"{a * 100:.0f}~{min(b, 1) * 100:.0f}%", "시료 수": int(m.sum()),
                        "평균 예측 확률(%)": float(pok0[m].mean() * 100), "실제 OK 비율(%)": float(ok_true[m].mean() * 100)})
    return {"rmse": rmse, "r2": r2, "sd": tm.std(axis=0), "pred": pred0, "pok": pok0, "ok_true": ok_true,
            "tp": tp, "tn": tn, "fp": fp, "fn": fn, "auc": auc, "acc": (tp + tn) / n,
            "ok_recall": tp / n1 if n1 else float("nan"), "ng_recall": tn / n0 if n0 else float("nan"),
            "calib": pd.DataFrame(cal)}


def max_rect(mask):
    """True 칸으로만 이루어진 가장 넓은 직사각형 (행 시작, 행 끝, 열 시작, 열 끝). 없으면 None."""
    nr, nc = mask.shape
    best_area, best = 0, None
    h = np.zeros(nc, dtype=int)
    for i in range(nr):
        h = np.where(mask[i], h + 1, 0)
        stack = []
        for j in range(nc + 1):
            cur = int(h[j]) if j < nc else 0
            start = j
            while stack and stack[-1][1] >= cur:
                s_idx, s_h = stack.pop()
                area = s_h * (j - s_idx)
                if area > best_area:
                    best_area, best = area, (i - s_h + 1, i, s_idx, j - 1)
                start = s_idx
            stack.append((start, cur))
    return best


@st.cache_data(show_spinner=False)
def condition_map(x, y, tm, lo, hi, quad, xlim, ylim, wx, wy, min_n, target, step=0.02):
    """허용 범위(xlim, ylim) 안의 격자 조건마다 OK 확률(모델)과 주변 실측 시료 수를 구하고,
    '시료 근거가 충분(min_n 이상)하면서 OK 확률 >= target' 인 칸으로만 이루어진 가장 넓은 직사각형을 찾는다."""
    x, y, tm = np.asarray(x, float), np.asarray(y, float), np.asarray(tm, float)
    lo, hi = np.asarray(lo, float), np.asarray(hi, float)
    model = fit_condition_model(x, y, tm, quad)
    gx = np.round(np.arange(np.ceil(xlim[0] / step - 1e-9) * step, xlim[1] + 1e-9, step), 6)
    gy = np.round(np.arange(np.ceil(ylim[0] / step - 1e-9) * step, ylim[1] + 1e-9, step), 6)
    GX, GY = np.meshgrid(gx, gy, indexing="ij")
    fx, fy = GX.ravel(), GY.ravel()
    P = pok_at(model, fx, fy, lo, hi)
    cnt = np.zeros(len(fx), dtype=int)
    for s in range(0, len(fx), 800):
        cnt[s:s + 800] = ((np.abs(fx[s:s + 800, None] - x[None, :]) <= wx + 1e-9) &
                          (np.abs(fy[s:s + 800, None] - y[None, :]) <= wy + 1e-9)).sum(axis=1)
    P, cnt = P.reshape(GX.shape), cnt.reshape(GX.shape)
    supported = cnt >= min_n
    ok_true = np.all((tm >= lo) & (tm <= hi), axis=1)

    rect = max_rect(supported & (P >= target))
    rect_info = None
    if rect is not None:
        i0, i1, j0, j1 = rect
        xr, yr = (float(gx[i0]), float(gx[i1])), (float(gy[j0]), float(gy[j1]))
        inside = (x >= xr[0]) & (x <= xr[1]) & (y >= yr[0]) & (y <= yr[1])
        rect_info = {"x": xr, "y": yr, "n_cells": int((i1 - i0 + 1) * (j1 - j0 + 1)), "p_min": float(P[i0:i1 + 1, j0:j1 + 1].min()),
                     "n_real": int(inside.sum()), "ok_real": int(ok_true[inside].sum())}

    # 근거가 충분한 칸 중 OK 확률이 높은 순 (서로 가까운 후보는 건너뛴다)
    order = np.lexsort((-cnt.ravel(), -P.ravel()))
    tops = []
    for o in order:
        if not supported.ravel()[o]:
            continue
        cx, cy = float(fx[o]), float(fy[o])
        if all(abs(cx - t["전장"]) > 2 * wx or abs(cy - t["일자-Y"]) > 2 * wy for t in tops):
            near = (np.abs(x - cx) <= wx) & (np.abs(y - cy) <= wy)
            tops.append({"전장": cx, "일자-Y": cy, "OK 확률(모델, %)": float(P.ravel()[o] * 100), "주변 실측 시료 수": int(near.sum()),
                         "주변 실측 OK 비율(%)": float(ok_true[near].mean() * 100) if near.any() else float("nan")})
        if len(tops) >= 5:
            break
    pred_top = predict_condition(model, [t["전장"] for t in tops], [t["일자-Y"] for t in tops]) if tops else np.empty((0, tm.shape[1]))
    return {"gx": gx, "gy": gy, "P": P, "cnt": cnt, "supported": supported, "rect": rect_info, "tops": pd.DataFrame(tops), "pred_top": pred_top,
            "step": step}


@st.cache_data(show_spinner=False)
def plan_adjustments(x, y, tm, lo, hi, quad, rmse, xlim, ylim, limx, limy, k, scales, step=0.01):
    """스펙을 벗어난(NG) 시료마다, 허용 범위와 변경 한도 안에서 '가장 적게 바꿔서' 결과 3개가 모두 스펙 안에 들어갈 것으로
    예측되는 (전장, 일자-Y) 를 찾는다. 시료마다 자기 고유의 오차(잔차)는 그대로 유지된다고 가정하고,
    스펙 경계에서 k × (교차검증 오차) 만큼 안쪽에 들어오도록 안전 여유를 둔다."""
    x, y, tm = np.asarray(x, float), np.asarray(y, float), np.asarray(tm, float)
    lo, hi, rmse = np.asarray(lo, float), np.asarray(hi, float), np.asarray(rmse, float)
    model = fit_condition_model(x, y, tm, quad)
    gx = np.round(np.arange(np.ceil(xlim[0] / step - 1e-9) * step, xlim[1] + 1e-9, step), 6)
    gy = np.round(np.arange(np.ceil(ylim[0] / step - 1e-9) * step, ylim[1] + 1e-9, step), 6)
    GX, GY = np.meshgrid(gx, gy, indexing="ij")
    GX, GY = GX.ravel(), GY.ravel()
    Pg = predict_condition(model, GX, GY)
    up_m = k * rmse
    lo_m = np.where(lo > 0, k * rmse, 0.0)  # 0 은 물리적 하한이라 안전 여유를 두지 않는다
    ok_true = np.all((tm >= lo) & (tm <= hi), axis=1)
    rows = []
    for i in np.flatnonzero(~ok_true):
        pred = Pg + model["resid"][i]
        feas = (np.abs(GX - x[i]) <= limx + 1e-9) & (np.abs(GY - y[i]) <= limy + 1e-9)
        for t in range(tm.shape[1]):
            feas &= (pred[:, t] >= lo[t] + lo_m[t]) & (pred[:, t] <= hi[t] - up_m[t])
        if feas.any():
            cand = np.flatnonzero(feas)
            d = np.hypot((GX[cand] - x[i]) / scales[0], (GY[cand] - y[i]) / scales[1])
            j = cand[int(np.argmin(d))]
            rows.append({"idx": int(i), "ok": True, "nx": float(GX[j]), "ny": float(GY[j]), **{f"p{t}": float(pred[j, t]) for t in range(tm.shape[1])}})
        else:
            rows.append({"idx": int(i), "ok": False, "nx": np.nan, "ny": np.nan, **{f"p{t}": np.nan for t in range(tm.shape[1])}})
    plan = pd.DataFrame(rows)
    plan["pok"] = np.nan
    good = plan["ok"].to_numpy()
    if good.any():
        plan.loc[good, "pok"] = pok_at(model, plan.loc[good, "nx"].to_numpy(), plan.loc[good, "ny"].to_numpy(), lo, hi) * 100
    return plan


def build_cond_xlsx(acc_df, cls_df, calib_df, top_df, rect_df, plan_df, params) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        acc_df.to_excel(w, sheet_name="예측 정확도", index=False)
        cls_df.to_excel(w, sheet_name="OK_NG 분류 정확도", index=False)
        calib_df.to_excel(w, sheet_name="확률 보정", index=False)
        if rect_df is not None:
            rect_df.to_excel(w, sheet_name="추천 작업 구간", index=False)
        top_df.to_excel(w, sheet_name="추천 조건 후보", index=False)
        plan_df.to_excel(w, sheet_name="NG 시료 조정 추천", index=False)
        pd.DataFrame(list(params.items()), columns=["조건", "값"]).to_excel(w, sheet_name="분석 조건", index=False)
        for ws in w.book.worksheets:
            for idx, col in enumerate(ws.iter_cols(min_row=1, max_row=1), start=1):
                ws.column_dimensions[get_column_letter(idx)].width = max(13, len(str(col[0].value or "")) + 4)
    return bio.getvalue()


def pok_heatmap(cm, min_n, rect=None, pts=None):
    gx, gy, P, cnt, step = cm["gx"], cm["gy"], cm["P"], cm["cnt"], cm["step"]
    gxx, gyy = np.meshgrid(gx, gy, indexing="ij")
    d = pd.DataFrame({"x1": gxx.ravel() - step / 2, "x2": gxx.ravel() + step / 2, "y1": gyy.ravel() - step / 2, "y2": gyy.ravel() + step / 2,
                      "p": P.ravel() * 100, "n": cnt.ravel(), "cx": gxx.ravel(), "cy": gyy.ravel()})
    sup, uns = d[d["n"] >= min_n], d[d["n"] < min_n]
    enc = dict(x=alt.X("x1:Q", title="전장", scale=alt.Scale(zero=False)), x2="x2:Q", y=alt.Y("y1:Q", title="일자-Y", scale=alt.Scale(zero=False)), y2="y2:Q")
    tip = [alt.Tooltip("cx:Q", format=".2f", title="전장"), alt.Tooltip("cy:Q", format=".2f", title="일자-Y"),
           alt.Tooltip("p:Q", format=".1f", title="OK 확률(모델, %)"), alt.Tooltip("n:Q", title="주변 실측 시료 수")]
    layers = []
    if len(uns):
        layers.append(alt.Chart(uns).mark_rect(color="#374151", opacity=0.35).encode(**enc, tooltip=tip))
    if len(sup):
        layers.append(alt.Chart(sup).mark_rect().encode(
            **enc, color=alt.Color("p:Q", title="OK 확률(%)", scale=alt.Scale(domain=[0, 50, 100], range=["#7f1d1d", "#f59e0b", "#10b981"])), tooltip=tip))
    if pts is not None:
        layers.append(alt.Chart(pts).mark_circle(size=14, opacity=0.5).encode(
            x="x:Q", y="y:Q", color=alt.Color("status:N", title="실측", scale=alt.Scale(domain=["OK", "NG"], range=["#ffffff", "#ef4444"]))))
    if rect is not None:
        layers.append(alt.Chart(pd.DataFrame({"x1": [rect["x"][0] - step / 2], "x2": [rect["x"][1] + step / 2],
                                              "y1": [rect["y"][0] - step / 2], "y2": [rect["y"][1] + step / 2]})).mark_rect(
            fillOpacity=0, stroke="#ffffff", strokeWidth=2.5).encode(x="x1:Q", x2="x2:Q", y="y1:Q", y2="y2:Q"))
    return _dark(alt.layer(*layers), 440)


def pred_actual_chart(actual, pred, label, hi_spec):
    d = pd.DataFrame({"a": np.asarray(actual, float), "p": np.asarray(pred, float)})
    lo_v, hi_v = float(min(d["a"].min(), d["p"].min())), float(max(d["a"].max(), d["p"].max()))
    pts = alt.Chart(d).mark_circle(size=26, opacity=0.5, color=ACCENT_BLUE).encode(
        x=alt.X("p:Q", title=f"예측 {label} (교차검증)", scale=alt.Scale(domain=[lo_v, hi_v])),
        y=alt.Y("a:Q", title=f"실측 {label}", scale=alt.Scale(domain=[lo_v, hi_v])),
        tooltip=[alt.Tooltip("p:Q", format=".3f", title="예측"), alt.Tooltip("a:Q", format=".3f", title="실측")])
    diag = alt.Chart(pd.DataFrame({"x": [lo_v, hi_v], "y": [lo_v, hi_v]})).mark_line(color="#ffffff", strokeDash=[6, 4]).encode(x="x:Q", y="y:Q")
    lim_v = alt.Chart(pd.DataFrame({"v": [hi_spec]})).mark_rule(color="#fb7185", strokeDash=[6, 4], strokeWidth=2).encode(y="v:Q")
    lim_h = alt.Chart(pd.DataFrame({"v": [hi_spec]})).mark_rule(color="#fb7185", strokeDash=[6, 4], strokeWidth=2).encode(x="v:Q")
    return _dark(alt.layer(pts, diag, lim_v, lim_h), 360)


def calib_chart(calib):
    d = calib.rename(columns={"평균 예측 확률(%)": "pred", "실제 OK 비율(%)": "act", "시료 수": "n"})
    pts = alt.Chart(d).mark_circle(color=SUCCESS, opacity=0.9).encode(
        x=alt.X("pred:Q", title="예측 OK 확률 (%)", scale=alt.Scale(domain=[0, 100])),
        y=alt.Y("act:Q", title="실제 OK 비율 (%)", scale=alt.Scale(domain=[0, 100])),
        size=alt.Size("n:Q", legend=None, scale=alt.Scale(range=[80, 900])),
        tooltip=[alt.Tooltip("예측 OK 확률 구간:N"), alt.Tooltip("n:Q", title="시료 수"), alt.Tooltip("pred:Q", format=".1f", title="평균 예측(%)"),
                 alt.Tooltip("act:Q", format=".1f", title="실제 OK(%)")])
    diag = alt.Chart(pd.DataFrame({"x": [0, 100], "y": [0, 100]})).mark_line(color="#ffffff", strokeDash=[6, 4]).encode(x="x:Q", y="y:Q")
    return _dark(alt.layer(pts, diag), 300)


def adjust_arrow_chart(cur, rec):
    """현재 조건(빨강) -> 추천 조건(초록) 이동 화살표(선)."""
    both = pd.concat([cur.assign(kind="현재(NG)"), rec.assign(kind="추천 조건")])
    seg = alt.Chart(rec.assign(x0=cur["x"].values, y0=cur["y"].values)).mark_rule(color="#9ca3af", opacity=0.55).encode(
        x="x0:Q", x2="x:Q", y="y0:Q", y2="y:Q")
    pts = alt.Chart(both).mark_circle(size=34, opacity=0.8).encode(
        x=alt.X("x:Q", title="전장", scale=alt.Scale(zero=False)), y=alt.Y("y:Q", title="일자-Y", scale=alt.Scale(zero=False)),
        color=alt.Color("kind:N", title=None, scale=alt.Scale(domain=["현재(NG)", "추천 조건"], range=[DANGER, SUCCESS])),
        tooltip=[alt.Tooltip("sample:N", title="시료"), alt.Tooltip("x:Q", format=".3f", title="전장"), alt.Tooltip("y:Q", format=".3f", title="일자-Y")])
    return _dark(alt.layer(seg, pts), 400)


# ============================================================
# 2. UI (원본과 동일한 로그인 → 사이드바 입력 → 상단 탭 구조)
# ============================================================
login_gate()

if "var_rules" not in st.session_state:
    st.session_state["var_rules"] = dict(DEFAULT_VARIABLE_RULES)
if "target_rules" not in st.session_state:
    st.session_state["target_rules"] = dict(DEFAULT_TARGET_RULES)
if "name_pattern" not in st.session_state:
    st.session_state["name_pattern"] = DEFAULT_NAME_PATTERN
if "end_rules" not in st.session_state:
    st.session_state["end_rules"] = dict(DEFAULT_END_RULES)


def _run_with_progress(key_prefix: str, work_fn):
    """진행 상황을 화면 정중앙의 고정 위치 박스(어두운 배경 포함)로 보여주면서 work_fn을
    실행한다.

    일부러 Streamlit 공식 st.dialog를 쓰지 않는다. st.dialog는 배경(바깥 영역)을
    클릭하면 자동으로 닫혀버리는 동작이 내장돼 있고, 이건 브라우저/Streamlit 내부 JS가
    하는 일이라 CSS(pointer-events 등)로 막으려는 시도가 두 번 다 실패했다
    (2026-09-20). 대신 BA-ROLLER 코드에서 검증된 방식대로, 진짜 모달 API 없이 이 함수가
    직접 전체화면 배경 div + 박스 div를 그린다 - 이 div들에는애초에 '클릭하면 닫기' 핸들러
    자체를 달지 않으므로, 바깥을 클릭해도 그냥 아무 일도 안 일어난다(닫히는 동작 자체가
    존재하지 않음). 박스는 남은 작업이 끝나고 우리가 명시적으로 지울 때만 사라진다.

    work_fn은 update_progress(현재_단계_라벨) 콜백 하나를 인자로 받는 함수여야 하고, 그
    안에서 실제 작업을 하다가 단계가 바뀔 때마다 update_progress(다음 라벨)를 불러
    진행률/문구를 갱신하면 된다 (전체 단계 수는 work_fn이 스스로 알고 있으므로,
    update_progress 호출을 '현재까지 몇 번째'로 세어 자동으로 %를 계산한다).

    끝나면 결과(성공 시 반환값, 실패 시 예외 메시지+트레이스백)를
    session_state[f"{key_prefix}_dialog_result"] / f"{key_prefix}_dialog_error"에 저장하고
    박스를 지운 뒤 rerun한다 - 그 다음 리런에서, 버튼을 그린 코드 바로 아래가 이 값을 읽어
    지금까지처럼 st.success/st.error로 보여주면 된다 (호출부는 예전 st.dialog 버전과 완전히
    동일하게 쓰면 되고 바꿀 필요가 없다)."""
    backdrop_cls = f"_bacapa_backdrop_{key_prefix}"
    box_cls = f"{backdrop_cls}_box"
    spin_cls = f"{backdrop_cls}_spin"
    # 주의: Streamlit의 st.markdown은 HTML을 넣기 전에 마크다운 파서를 먼저 거치는데,
    # 마크다운 문법상 '줄 앞에 공백 4칸 이상 들여쓰기'는 코드블록으로 해석돼서 HTML이
    # 아니라 글자 그대로("<div class=...>") 화면에 찍혀버린다 (2026-09-20 실제 발생 -
    # Python 코드 스타일대로 들여써서 넣은 게 원인). 그래서 여기 들어가는 HTML/CSS 문자열은
    # 반드시 줄마다 앞쪽 공백을 없앤 채로(한 줄씩 이어붙이는 방식으로) 만든다.
    st.markdown(
        f'<style>'
        f'.{backdrop_cls} {{ position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; '
        f'background: rgba(8,8,8,0.82); z-index: 99998; }}'
        f'.{box_cls} {{ position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); '
        f'z-index: 99999; background: #131d34; border-radius: 14px; '
        f'box-shadow: 0 20px 60px rgba(0,0,0,0.7); width: 340px; max-width: 92vw; '
        f'box-sizing: border-box; padding: 22px 24px; }}'
        f'@keyframes {spin_cls} {{ to {{ transform: rotate(360deg); }} }}'
        f'.{spin_cls} {{ display: inline-block; width: 15px; height: 15px; margin-right: 8px; '
        f'border: 3px solid rgba(255,255,255,0.25); border-top-color: {ACCENT_BLUE}; '
        f'border-radius: 50%; animation: {spin_cls} 0.8s linear infinite; '
        f'vertical-align: middle; flex-shrink: 0; }}'
        f'</style>'
        f'<div class="{backdrop_cls}"></div>',
        unsafe_allow_html=True,
    )

    box_ph = st.empty()

    def render(label: str, pct: int):
        box_ph.markdown(
            f'<div class="{box_cls}">'
            f'<div style="font-weight:700;color:{TEXT_MAIN};font-size:1rem;margin-bottom:10px;">처리 중</div>'
            f'<div style="display:flex;align-items:center;color:{TEXT_MAIN};font-size:0.9rem;">'
            f'<span class="{spin_cls}"></span>{label}</div>'
            f'<div style="background:rgba(255,255,255,0.12);border-radius:6px;height:8px;overflow:hidden;margin-top:14px;">'
            f'<div style="background:{ACCENT_BLUE};height:100%;width:{pct}%;transition:width .3s;"></div></div>'
            f'<div style="text-align:right;font-size:0.78rem;color:{TEXT_MUTED};margin-top:4px;">{pct}%</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    calls = {"n": 0}

    def update_progress(label: str):
        calls["n"] += 1
        # 총 단계 수를 미리 알 수 없으니, 부드럽게 100%에 가까워지도록 점근적으로 채운다.
        pct = min(95, int(100 * (1 - 1 / (1 + calls["n"]))))
        render(label, pct)

    update_progress("처리를 시작합니다...")
    try:
        result = work_fn(update_progress)
        render("완료되었습니다", 100)
        st.session_state[f"{key_prefix}_dialog_result"] = result
        st.session_state[f"{key_prefix}_dialog_error"] = None
    except Exception as e:
        import traceback
        st.session_state[f"{key_prefix}_dialog_result"] = None
        st.session_state[f"{key_prefix}_dialog_error"] = (str(e), traceback.format_exc())
    box_ph.empty()
    st.rerun()


def render_header():
    res = st.session_state.get("mm_result")
    status_txt = "분석 완료" if res else "대기 중"
    status_color = SUCCESS if res else TEXT_MUTED
    pass_txt = f"{len(res['subset'])}개 / {res['info']['n_samples']}개" if res else "-"
    st.markdown(
        f"""<div class='glass-card' style='padding:18px 22px;margin-bottom:14px;display:flex;align-items:center;gap:18px;'>
            <img src='data:image/png;base64,{ILJIN_LOGO_B64}' style='height:20px;flex-shrink:0;'/>
            <div style='width:1px;height:52px;background:{BORDER_STRONG};flex-shrink:0;'></div>
            <div>
                <div style='color:{ACCENT_BLUE};font-size:2.04rem;font-weight:700;letter-spacing:0.02em;margin-bottom:4px;line-height:1.2;'>WELDING ARM 측정 데이터 분석</div>
                <div style='display:flex;align-items:baseline;gap:10px;margin-bottom:2px;'>
                    <span style='color:{TEXT_MAIN};font-size:1.3rem;font-weight:700;border-bottom:2px solid {ACCENT_BLUE};padding-bottom:2px;'>타겟 스펙 만족 시료의 전장 · 일자-Y Min / Max 추출</span>
                    <span style='color:{TEXT_MUTED};font-size:0.85rem;'>V1.0</span>
                </div>
                <div style='color:{TEXT_MUTED};font-size:0.85rem;'>
                    상태: <span style='color:{status_color};font-weight:600;'>{status_txt}</span>
                    &nbsp;·&nbsp; 스펙 합격 시료: <b style='color:{TEXT_MAIN};'>{pass_txt}</b>
                </div>
            </div>
        </div>""",
        unsafe_allow_html=True,
    )


render_header()

# ------------------------------------------------------------------
# 사이드바: 측정 데이터 업로드 (분석 조건 입력/실행/결과는 각 분석 탭에서)
# ------------------------------------------------------------------
with st.sidebar:
    st.markdown(f"<div style='margin-bottom:6px;'><img src='data:image/png;base64,{ILJIN_LOGO_B64}' style='height:31px;'/></div>", unsafe_allow_html=True)
    st.caption("용접 암 측정 데이터 업로드 → 분석은 위쪽 탭에서")
    st.markdown("---")

    with st.expander("STEP 1 · 측정 데이터 업로드", expanded=True):
        st.markdown('<div class="block-title">측정 데이터 엑셀 업로드 — 파일명 예: DATA-KW</div>', unsafe_allow_html=True)
        data_file = st.file_uploader("측정 데이터 엑셀 업로드", type=["xlsx"], label_visibility="collapsed", key="data_uploader")

        data_bytes, sheet_name = None, None
        if data_file is not None:
            data_bytes = data_file.getvalue()
            try:
                sheets = _list_sheets(data_bytes)
                sheet_name = sheets[0] if len(sheets) == 1 else st.selectbox("시트 선택", sheets, key="data_sheet")
            except Exception as e:
                st.error(f"엑셀을 여는 중 오류가 발생했습니다 (열기 암호가 걸린 파일은 열 수 없습니다): {e}")
                data_bytes = None

        if data_bytes is None:
            st.info("측정 데이터 엑셀을 업로드하면 위쪽 분석 탭에서 실행할 수 있습니다.")
        else:
            st.caption("업로드 완료 — 분석 조건 입력과 실행은 위쪽 Tab 1에서 합니다.")

    st.markdown("---")
    if st.button("로그아웃"):
        st.session_state["authenticated"] = False
        st.rerun()


# ------------------------------------------------------------------
# 상단 탭
# ------------------------------------------------------------------
if st.session_state.get("is_owner"):
    NAV_ITEMS = ["Tab 1. 스펙 만족 시료 Min/Max 분석", "Tab 2. 조건 찾기", "Tab 3. 시료 확인", "Tab 4. Go/No-Go 게이지 설계", "Tab 5. 품질 분석",
                 "처리 이력 · 통계", "분석 설정 관리", "참고 / 한계 사항"]
    tab_main, tab_cond, tab_samples, tab_gauge, tab_analysis, tab_history, tab_sync, tab_help = st.tabs(NAV_ITEMS)
else:
    tab_main, tab_cond, tab_samples, tab_gauge, tab_analysis = st.tabs(
        ["Tab 1. 스펙 만족 시료 Min/Max 분석", "Tab 2. 조건 찾기", "Tab 3. 시료 확인", "Tab 4. Go/No-Go 게이지 설계", "Tab 5. 품질 분석"])
    tab_history = tab_sync = tab_help = None


def _kpi(label, value, sub="", color=None):
    style = f' style="color:{color}"' if color else ""
    sub_html = f'<p class="kpi-label" style="margin-top:4px;">{sub}</p>' if sub else ""
    return f'<div class="kpi-card"><p class="kpi-label">{label}</p><p class="kpi-value"{style}>{value}</p>{sub_html}</div>'


def _cols_table(cols, info=None):
    rows = [
        {"구분": v["kind"], "항목": k, "인식한 머리글": v["header"], "값 열": get_column_letter(v["value_col"] + 1),
         "편차 열": get_column_letter(v["dev_col"] + 1) if v["dev_col"] is not None else "-"}
        for k, v in cols.items()
    ]
    for side, d in ((info or {}).get("end_cols") or {}).items():
        rows.append({"구분": "끝단", "항목": f"{side} 원주 {len(d['cols'])}점",
                     "인식한 머리글": f"{d['headers'][0]} ~ {d['headers'][-1]}",
                     "값 열": ", ".join(get_column_letter(c + 1) for c in d["cols"]), "편차 열": "-"})
    return pd.DataFrame(rows)


with tab_main:
    st.markdown('<div class="block-title">분석 조건 · 타겟 항목 스펙 (3개 모두 만족)</div>', unsafe_allow_html=True)
    specs = {}
    _tcols = st.columns(len(st.session_state["target_rules"]))
    for _holder, _key in zip(_tcols, st.session_state["target_rules"]):
        with _holder:
            _d_lo, _d_hi = DEFAULT_TARGET_SPECS.get(_key, (0.0, 1.0))
            st.markdown(f"<div style='font-size:0.9rem;color:{TEXT_MAIN};font-weight:600;'>{_key}</div>", unsafe_allow_html=True)
            _c1, _c2 = st.columns(2)
            _lo = _c1.number_input("하한", value=float(_d_lo), step=0.1, format="%.3f", key=f"spec_lo_{_key}")
            _hi = _c2.number_input("상한", value=float(_d_hi), step=0.1, format="%.3f", key=f"spec_hi_{_key}")
            specs[_key] = (_lo, _hi)
    st.caption("하한/상한 경계값도 합격에 포함됩니다. 변수 항목은 전장, 일자-Y 입니다.")

    if data_bytes is None:
        st.info("왼쪽 사이드바 STEP 1에서 측정 데이터 엑셀을 먼저 업로드해주세요.")

    if st.button("Min/Max 분석 실행", type="primary", disabled=data_bytes is None, key="spec_run"):
        _var_rules = dict(st.session_state["var_rules"])
        _target_rules = dict(st.session_state["target_rules"])
        _name_pat = st.session_state["name_pattern"]
        _end_rules = dict(st.session_state["end_rules"])
        _file_name = data_file.name
        _specs = dict(specs)

        def _analyze_work(update_progress):
            update_progress("엑셀 파일을 읽는 중...")
            raw = _load_raw_sheet(data_bytes, sheet_name)
            update_progress("타겟/변수 열을 찾고 시료 행만 추리는 중...")
            df, cols, info = prepare_dataset(raw, _var_rules, _target_rules, _name_pat, _end_rules)
            update_progress("3개 타겟 스펙을 모두 만족하는 시료를 고르는 중...")
            summary, subset, status_df = analyze_spec(df, _specs, list(_var_rules))
            end_summary = end_extent_summary(subset) if (info.get("has_end") and not subset.empty) else None
            update_progress("참고용 합격 보장 범위(불합격 0개)를 찾는 중...")
            _vk = list(_var_rules)
            _flags = pass_matrix(df, _specs)["전체"].values
            win = find_pass_window(df, _flags, _vk[0], _vk[1]) if len(_vk) >= 2 else None
            update_progress("허용 불합격률별 범위를 비교하는 중...")
            trade = window_tradeoff(df, _flags, _vk[0], _vk[1]) if win is not None else []
            trade_end = window_tradeoff(df, _flags, "양 끝단 포함 전장", _vk[1]) if (win is not None and "양 끝단 포함 전장" in df.columns and df["양 끝단 포함 전장"].notna().all()) else []

            update_progress("결과 파일 만드는 중...")
            spec_txt = " / ".join(f"{k} {min(v):g}~{max(v):g}" for k, v in _specs.items())
            params = {
                "데이터 파일": _file_name, "시트": sheet_name, "타겟 스펙(경계 포함)": spec_txt,
                "변수 항목": ", ".join(_var_rules), "유효 시료 수": info["n_samples"], "스펙 합격 시료 수": len(subset),
            }
            st.session_state["mm_result"] = {
                "summary": summary, "subset": subset, "status": status_df, "specs": _specs, "spec_txt": spec_txt,
                "var_keys": list(_var_rules), "cols": cols, "info": info, "file": _file_name, "sheet": sheet_name,
                "df": df, "nominals": {k: _last_number(cols[k]["header"]) for k in _var_rules}, "window": win, "tradeoff": trade, "tradeoff_end": trade_end, "end_summary": end_summary,
                "xlsx": build_result_xlsx(summary, status_df, subset, params, win, end_summary) if not subset.empty else None,
            }

            update_progress("처리 이력 기록하는 중...")
            hist = st.session_state.setdefault("process_history", [])
            rec = {"시각": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "구분": "타겟 스펙 합격",
                   "데이터 파일": _file_name, "조건": spec_txt, "시료 수": len(subset)}
            if not summary.empty:
                for key in _var_rules:
                    r = summary[(summary["항목"] == key) & (summary["구분"] == "측정값")].iloc[0]
                    rec[f"{key} 최종 Min ~ Max"] = f"{r['Min']:.4f} ~ {r['Max']:.4f}"
            if end_summary is not None and not end_summary.empty:
                _e = end_summary[end_summary["항목"] == "양 끝단 포함 전장"].iloc[0]
                rec["양 끝단 포함 전장 Min ~ Max"] = f"{_e['Min']:.4f} ~ {_e['Max']:.4f}"
            if win is not None:
                rec[f"{win['xkey']} 보장범위(참고)"] = f"{win['x'][0]:.4f} ~ {win['x'][1]:.4f}"
                rec[f"{win['ykey']} 보장범위(참고)"] = f"{win['y'][0]:.4f} ~ {win['y'][1]:.4f}"
            hist.insert(0, rec)
            st.session_state.setdefault("spec_counter", Counter()).update([spec_txt])
            return {"n_in": len(subset), "spec_txt": spec_txt}

        _run_with_progress("analyze", _analyze_work)

    if st.session_state.get("analyze_dialog_result") is not None or \
            st.session_state.get("analyze_dialog_error") is not None:
        _err = st.session_state.pop("analyze_dialog_error", None)
        _res = st.session_state.pop("analyze_dialog_result", None)
        if _err:
            _msg, _tb = _err
            st.error(f"분석 중 오류가 발생했습니다: {_msg}")
            with st.expander("상세 오류"):
                st.code(_tb)
        elif _res["n_in"] == 0:
            st.warning("3개 타겟 스펙을 모두 만족하는 시료가 없습니다. 스펙 값을 확인해주세요.")
        else:
            st.success(f"3개 타겟 스펙을 모두 만족하는 시료 {_res['n_in']}건으로 전장 · 일자-Y 최종 Min/Max를 계산했습니다. 아래 결과를 확인하세요.")

    st.markdown("<hr>", unsafe_allow_html=True)
    res = st.session_state.get("mm_result")
    if res is None:
        st.info("아직 분석 결과가 없습니다. 위에서 타겟 스펙을 확인하고 'Min/Max 분석 실행'을 눌러주세요.")
    else:
        summary, subset, status_df, info = res["summary"], res["subset"], res["status"], res["info"]
        win = res.get("window")
        n_all, n_in = info["n_samples"], len(subset)
        pct = (n_in / n_all * 100) if n_all else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(_kpi("합격 조건", "3개 타겟 동시 만족", "단축·장축 위치도, 단축 윤곽도"), unsafe_allow_html=True)
        k2.markdown(_kpi("유효 시료 수", f"{n_all}개", "엑셀 시료 행 기준"), unsafe_allow_html=True)
        k3.markdown(_kpi("스펙 합격(OK) 시료 수", f"{n_in}개", "3개 타겟 스펙 모두 만족", color=SUCCESS if n_in else WARNING), unsafe_allow_html=True)
        k4.markdown(_kpi("합격률", f"{pct:.1f}%", "합격 시료 ÷ 유효 시료"), unsafe_allow_html=True)

        notes = []
        if info["n_excluded_non_sample"]:
            notes.append(f"MIN/MAX/AVG 등 요약·구분 행 {info['n_excluded_non_sample']}개는 시료가 아니므로 제외")
        if info["n_incomplete"]:
            notes.append(f"값이 비어 있는 시료 행 {info['n_incomplete']}개 제외")
        if notes:
            st.caption(" · ".join(notes))

        st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
        st.markdown('<div class="block-title">① 합격(OK) 시료의 최종 Min / Max — 전장 · 일자-Y</div>', unsafe_allow_html=True)
        if summary.empty:
            st.warning("3개 타겟 스펙을 모두 만족하는 시료가 없어 Min/Max를 계산하지 못했습니다.")
        else:
            vals = summary[summary["구분"] == "측정값"].set_index("항목")
            card_cols = st.columns(2 * len(res["var_keys"]))
            idx = 0
            for key in res["var_keys"]:
                for kind in ("Min", "Max"):
                    r = vals.loc[key]
                    card_cols[idx].markdown(
                        _kpi(f"{key} {kind}", f"{r[kind]:.4f}", f"{r[f'{kind} 시료']} (엑셀 {r[f'{kind} 엑셀 행']}행)"),
                        unsafe_allow_html=True,
                    )
                    idx += 1
            st.markdown("<div style='margin-top:12px;'></div>", unsafe_allow_html=True)
            _xk, _yk = res["var_keys"][0], res["var_keys"][1]
            st.success(
                f"3개 스펙을 모두 만족한 시료 {n_in}개를 하나씩 모아, 그 중 가장 작은 값과 가장 큰 값을 최종 Min/Max로 정했습니다. "
                f"따라서 합격 시료 {n_in}개는 전부 이 범위 안에 들어옵니다 "
                f"({_xk} {vals.loc[_xk, 'Min']:.4f} ~ {vals.loc[_xk, 'Max']:.4f}, {_yk} {vals.loc[_yk, 'Min']:.4f} ~ {vals.loc[_yk, 'Max']:.4f}).")
            _flags = pass_matrix(res["df"], res["specs"])["전체"].values
            _inbox = res["df"][_xk].between(vals.loc[_xk, "Min"], vals.loc[_xk, "Max"]) & res["df"][_yk].between(vals.loc[_yk, "Min"], vals.loc[_yk, "Max"])
            st.caption(
                f"참고: 반대로, 이 두 범위에 동시에 들어오는 시료는 {int(_inbox.sum())}개이고 그 중 {int((~_flags[_inbox.values]).sum())}개는 스펙을 만족하지 못한 시료입니다. "
                "즉 이 범위는 '합격 시료가 존재했던 폭'이며, 범위 안이라고 해서 합격이 보장되지는 않습니다 (보장되는 범위는 아래 ④ 참고).")
            num_cfg = {c: st.column_config.NumberColumn(format="%.4f") for c in ("Min", "Max", "Max-Min")}
            st.dataframe(summary, use_container_width=True, hide_index=True, column_config=num_cfg)

        st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
        st.markdown('<div class="block-title">② 양 끝단(파이프 기울기)을 포함한 전장 Min / Max — 지그 전장 치수용</div>', unsafe_allow_html=True)
        esum = res.get("end_summary")
        if esum is None or esum.empty:
            st.info("엑셀에서 Y-값_단1~8 / Y-값_장1~8 열을 찾지 못해 계산하지 않았습니다 (분석 설정 관리 탭에서 끝단 열 이름 규칙을 확인할 수 있습니다).")
        else:
            ev = esum.set_index("항목")
            ec = st.columns(4)
            _lr = ev.loc["양 끝단 포함 전장"]
            ec[0].markdown(_kpi("포함 전장 Min", f"{_lr['Min']:.4f}", f"{_lr['Min 시료']} (엑셀 {_lr['Min 엑셀 행']}행)"), unsafe_allow_html=True)
            ec[1].markdown(_kpi("포함 전장 Max", f"{_lr['Max']:.4f}", f"{_lr['Max 시료']} (엑셀 {_lr['Max 엑셀 행']}행)"), unsafe_allow_html=True)
            if "기울기 증가량" in ev.index:
                _gr = ev.loc["기울기 증가량"]
                ec[2].markdown(_kpi("기울기로 늘어난 길이 (평균)", f"+{_gr['평균']:.4f}", "포함 전장 − 전장(중심 기준)"), unsafe_allow_html=True)
                ec[3].markdown(_kpi("기울기로 늘어난 길이 (최대)", f"+{_gr['Max']:.4f}", f"{_gr['Max 시료']} (엑셀 {_gr['Max 엑셀 행']}행)"), unsafe_allow_html=True)
            st.markdown("<div style='margin-top:12px;'></div>", unsafe_allow_html=True)
            _w = res["info"].get("end_cols", {})
            _nl, _nr = len(_w.get("왼쪽 끝단", {}).get("cols", [])), len(_w.get("오른쪽 끝단", {}).get("cols", []))
            st.success(
                f"합격(OK) 시료 {n_in}개 각각에서, 왼쪽 끝단 원주 {_nl}점(Y-값_단) 중 가장 왼쪽 값과 오른쪽 끝단 원주 {_nr}점(Y-값_장) 중 가장 오른쪽 값 사이의 거리"
                f"(= 전장 + 양쪽 끝단이 기울어서 튀어나온 길이)를 구해, 그 최소 {_lr['Min']:.4f} ~ 최대 {_lr['Max']:.4f} 로 정했습니다. "
                f"가장 긴 합격 제품이 들어가려면 지그의 전장 방향 치수는 {_lr['Max']:.4f} 이상을 포함해야 합니다.")
            if res["info"].get("n_end_incomplete"):
                st.caption(f"끝단 측정점이 일부 비어 있는 시료 {res['info']['n_end_incomplete']}개는 있는 점만으로 계산했습니다.")
            if len(subset) >= 3:
                st.altair_chart(hist_chart(subset["양 끝단 포함 전장"], "양 끝단 포함 전장 (합격 시료)", [("Min", float(_lr["Min"])), ("Max", float(_lr["Max"]))]),
                                use_container_width=True, theme=None)
            ecfg = {c: st.column_config.NumberColumn(format="%.4f") for c in ("Min", "Max", "평균", "Max-Min")}
            st.dataframe(esum, use_container_width=True, hide_index=True, column_config=ecfg)
            st.caption("양 끝단 포함 전장 = 오른쪽 최외측 Y(가장 큰 값) − 왼쪽 최외측 Y(가장 작은 값). '안쪽 전장'은 양 끝단이 가장 가까워지는 길이(오른쪽 최내측 − 왼쪽 최내측)로, "
                       "지그에 걸림(스토퍼)이 있을 때 참고하세요. 지그 구조에 따라 필요한 치수가 달라질 수 있습니다.")

        st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
        st.markdown('<div class="block-title">③ 타겟 항목별 합격 현황</div>', unsafe_allow_html=True)
        st.caption("'단독 합격'은 그 항목 하나만 봤을 때의 합격 수입니다.")
        stat_cfg = {c: st.column_config.NumberColumn(format="%.4f") for c in ("전체 실측 Min", "전체 실측 Max")}
        stat_cfg["단독 합격률(%)"] = st.column_config.NumberColumn(format="%.1f")
        st.dataframe(status_df, use_container_width=True, hide_index=True, column_config=stat_cfg)

        with st.expander("④ 참고 · 이 범위 안이면 반드시 합격인 더 좁은 범위 (불합격 0개, Go/No-Go 게이지용)", expanded=False):
            if win is None:
                st.warning("3개 타겟 스펙을 모두 만족하는 시료가 없어 범위를 찾지 못했습니다.")
            else:
                xk, yk = win["xkey"], win["ykey"]
                st.caption(
                    f"위 ① 의 Min/Max 는 합격 시료의 양 끝값이라 그 안에 불합격 시료가 섞여 있습니다. 여기서는 {xk} 범위와 {yk} 범위에 **동시에** 들어오면 "
                    "불합격 시료가 하나도 없도록, 불합격이 포함되지 않는 직사각형 범위 중 합격 시료를 가장 많이 포함하는 범위를 찾았습니다. "
                    "이번 데이터에서 실제 측정된 시료 기준이며, 시료가 추가되면 달라질 수 있습니다.")
                wc = st.columns(4)
                for holder, label, val in ((wc[0], f"{xk} Min", win["x"][0]), (wc[1], f"{xk} Max", win["x"][1]),
                                           (wc[2], f"{yk} Min", win["y"][0]), (wc[3], f"{yk} Max", win["y"][1])):
                    holder.markdown(_kpi(label, f"{val:.4f}"), unsafe_allow_html=True)
                st.markdown("<div style='margin-top:12px;'></div>", unsafe_allow_html=True)
                if win["n_fail_in"] == 0:
                    st.success(
                        f"{xk} {win['x'][0]:.4f} ~ {win['x'][1]:.4f} 이면서 {yk} {win['y'][0]:.4f} ~ {win['y'][1]:.4f} 인 시료는 "
                        f"{win['n_in']}개이고, 전부 3개 타겟 스펙을 만족합니다 (불합격 0개). "
                        f"합격 시료 {win['n_pass']}개 중 {win['n_pass_in']}개({win['n_pass_in'] / win['n_pass'] * 100:.1f}%)를 포함합니다.")
                else:
                    st.warning(f"범위 안에 불합격 시료가 {win['n_fail_in']}개 섞여 있습니다. 데이터를 확인해주세요.")
                st.altair_chart(window_scatter_chart(res["df"], pass_matrix(res["df"], res["specs"])["전체"].values, win),
                                use_container_width=True, theme=None)
                st.caption("흰색 사각형이 이 범위입니다. 초록 점은 합격, 빨간 점은 불합격 시료이며, 사각형 안에는 빨간 점이 없습니다.")

        if res.get("xlsx"):
            st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
            st.download_button(
                "스펙만족_범위_결과.xlsx 다운로드",
                data=res["xlsx"],
                file_name="스펙만족_범위_결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

with tab_samples:
    _r = st.session_state.get("mm_result")
    if _r is None:
        st.info("아직 분석 결과가 없습니다. Tab 1에서 타겟 스펙을 확인하고 'Min/Max 분석 실행'을 눌러주세요.")
    else:
        st.markdown('<div class="block-title">① 스펙 합격 시료 전체 (전장 오름차순)</div>', unsafe_allow_html=True)
        if _r["subset"].empty:
            st.warning("3개 타겟 스펙을 모두 만족하는 시료가 없습니다.")
        else:
            st.caption(f"3개 타겟 스펙을 모두 만족하는 시료 {len(_r['subset'])}개 (유효 시료 {_r['info']['n_samples']}개 중) · 조건: {_r['spec_txt']}")
            st.dataframe(_r["subset"], use_container_width=True, hide_index=True)

        st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
        st.markdown('<div class="block-title">② 인식된 열 정보</div>', unsafe_allow_html=True)
        st.caption("엑셀 머리글에서 찾은 변수(전장, 일자-Y)와 타겟(위치도, 윤곽도) 항목의 열 위치입니다.")
        st.dataframe(_cols_table(_r["cols"], _r["info"]), use_container_width=True, hide_index=True)


def _analysis_result():
    r = st.session_state.get("mm_result")
    if r is None or "df" not in r:
        st.info("아직 분석 결과가 없습니다. Tab 1에서 타겟 스펙을 확인하고 'Min/Max 분석 실행'을 눌러주세요.")
        return None
    return r


def _analysis_items(r):
    return list(r["var_keys"]) + list(r["specs"].keys())


def _pick_dataset(r, key):
    choice = st.radio("분석 대상", ["전체 유효 시료", "스펙 합격 시료"], horizontal=True, key=key)
    return (r["df"] if choice == "전체 유효 시료" else r["subset"]), choice


def _spec_lines_for(r, k):
    """분포/관리도 차트에 그릴 기준선. 타겟은 실행 때 쓴 스펙(상한), 전장·일자-Y 는 기준값만."""
    lines = []
    if k in r["specs"]:
        lo, hi = sorted(r["specs"][k])
        if lo > 0:
            lines.append(("LSL", lo))
        lines.append(("USL", hi))
    else:  # 전장·일자-Y: 규격(공차)이 없는 관리 변수 -> 머리글에 적힌 기준값만 참고선으로 표시
        nom = (r.get("nominals") or {}).get(k)
        if nom is not None:
            lines.append(("기준값", nom))
    return lines


# ---------------- Tab 2. 조건 찾기 ----------------
with tab_cond:
    r = _analysis_result()
    if r is not None:
        dfc, specs = r["df"], r["specs"]
        xk, yk = r["var_keys"][0], r["var_keys"][1]
        xmode = "c"
        if "양 끝단 포함 전장" in dfc.columns and dfc["양 끝단 포함 전장"].notna().all():
            _xsel = st.radio("전장 기준", ["중심 기준 전장", "양 끝단 포함 전장 (지그 치수용)"], horizontal=True, key="cond_xbasis",
                             help="양 끝단 포함 전장 = 오른쪽 끝단 최외측 Y − 왼쪽 끝단 최외측 Y (파이프 기울기로 튀어나온 길이 포함). 지그는 이 값을 기준으로 만들어야 제품이 들어갑니다.")
            if _xsel.startswith("양 끝단"):
                xk, xmode = "양 끝단 포함 전장", "e"
        tg = list(specs)
        xa, ya = dfc[xk].to_numpy(dtype=float), dfc[yk].to_numpy(dtype=float)
        tm = dfc[tg].to_numpy(dtype=float)
        lo = np.array([min(specs[k]) for k in tg])
        hi = np.array([max(specs[k]) for k in tg])

        st.caption(
            f"{xk}·{yk} 를 입력 변수, {', '.join(tg)} 를 결과값으로 보고, 이번 데이터의 경향으로 '결과가 3개 스펙을 모두 만족하려면 입력을 어떻게 맞춰야 하는지'를 찾습니다. "
            "① 예측이 얼마나 믿을 만한지 확인하고, ② 스펙을 만족할 가능성이 높은 입력 조건(작업 구간)을 찾고, ③ 현재 NG 시료의 입력을 얼마나 바꾸면 되는지 추정합니다.")
        mdl = st.radio("예측 모델", ["2차 모델 (곡률 · 상호작용 포함)", "선형 모델"], horizontal=True, key="cond_model")
        quad = mdl.startswith("2차")
        cv = cv_condition_model(xa, ya, tm, lo, hi, quad)
        cv_alt = cv_condition_model(xa, ya, tm, lo, hi, not quad)

        # ---------------- ① 예측 정확도 ----------------
        st.markdown('<div class="block-title">① 예측 정확도 (학습에 쓰지 않은 시료로 검증한 값)</div>', unsafe_allow_html=True)
        alt_name = "선형" if quad else "2차"
        acc_rows = []
        for t, k in enumerate(tg):
            acc_rows.append({"결과": k, "스펙 상한": hi[t], "원래 표준편차": cv["sd"][t], "예측 오차 RMSE": cv["rmse"][t],
                             "RMSE ÷ 스펙 폭(%)": cv["rmse"][t] / (hi[t] - lo[t]) * 100, "설명력 R²": cv["r2"][t],
                             f"{alt_name} 모델 RMSE (비교)": cv_alt["rmse"][t]})
        acc_df = pd.DataFrame(acc_rows)
        acc_cfg = {c: st.column_config.NumberColumn(format="%.3f") for c in ("스펙 상한", "원래 표준편차", "예측 오차 RMSE", f"{alt_name} 모델 RMSE (비교)")}
        acc_cfg["RMSE ÷ 스펙 폭(%)"] = st.column_config.NumberColumn(format="%.1f")
        acc_cfg["설명력 R²"] = st.column_config.NumberColumn(format="%.2f")
        st.dataframe(acc_df, use_container_width=True, hide_index=True, column_config=acc_cfg)
        st.caption("예측 오차 RMSE 는 새 시료의 결과값을 예측할 때의 평균적인 오차이고, 스펙 폭에 비해 작을수록 믿을 만합니다. "
                   "설명력 R² 는 입력 두 개로 설명되는 결과 변동의 비율이며, 나머지는 이 데이터에 없는 다른 요인 때문입니다. "
                   "데이터를 5등분해서 4/5 로 학습하고 나머지 1/5 로 평가하는 것을 반복한 값입니다.")

        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(_kpi("OK/NG 판정 정확도", f"{cv['acc'] * 100:.1f}%", f"전체 {len(xa)}개 중 맞춘 비율"), unsafe_allow_html=True)
        k2.markdown(_kpi("OK 를 OK 로 맞힘", f"{cv['ok_recall'] * 100:.1f}%", f"실제 OK {cv['tp'] + cv['fn']}개 중 {cv['tp']}개"), unsafe_allow_html=True)
        k3.markdown(_kpi("NG 를 NG 로 맞힘", f"{cv['ng_recall'] * 100:.1f}%", f"실제 NG {cv['tn'] + cv['fp']}개 중 {cv['tn']}개",
                         color=WARNING if cv["ng_recall"] < 0.6 else SUCCESS), unsafe_allow_html=True)
        k4.markdown(_kpi("구분 능력 (AUC)", f"{cv['auc']:.3f}", "0.5=못 가름 · 1.0=완벽"), unsafe_allow_html=True)
        cm_df = pd.DataFrame({"": ["실제 OK", "실제 NG"], "예측 OK": [cv["tp"], cv["fp"]], "예측 NG": [cv["fn"], cv["tn"]]})
        c_l, c_r = st.columns([1, 1])
        with c_l:
            st.markdown('<div style="margin-top:12px;"></div>', unsafe_allow_html=True)
            st.dataframe(cm_df, use_container_width=True, hide_index=True)
            st.caption("예측 OK 확률이 50% 이상이면 OK 로 예측한 결과입니다. NG 를 NG 로 맞히는 비율이 낮으면, 입력만으로는 NG 를 걸러내기 어렵다는 뜻입니다 "
                       "(다른 요인이 NG 를 만들기 때문). 이 탭은 'OK 가능성이 높은 입력 구간'을 찾는 용도입니다.")
        with c_r:
            if len(cv["calib"]):
                st.altair_chart(calib_chart(cv["calib"]), use_container_width=True, theme=None)
        if len(cv["calib"]):
            ccfg = {c: st.column_config.NumberColumn(format="%.1f") for c in ("평균 예측 확률(%)", "실제 OK 비율(%)")}
            st.dataframe(cv["calib"], use_container_width=True, hide_index=True, column_config=ccfg)
            st.caption("'예측 OK 확률'이 높은 시료일수록 실제 OK 비율도 높아지는지 확인하는 표입니다 (점이 대각선에 가까울수록 확률을 믿을 수 있음).")
        sel_t = st.selectbox("예측 vs 실측 그래프 항목", tg, index=tg.index("단축 윤곽도") if "단축 윤곽도" in tg else 0, key="cond_pa_item")
        ti = tg.index(sel_t)
        st.altair_chart(pred_actual_chart(tm[:, ti], cv["pred"][:, ti], sel_t, float(hi[ti])), use_container_width=True, theme=None)
        st.caption("점이 흰 대각선에 가까울수록 예측이 정확합니다. 분홍 점선은 스펙 상한입니다.")

        st.markdown("<hr>", unsafe_allow_html=True)

        # ---------------- ② 추천 조건 ----------------
        st.markdown('<div class="block-title">② 3개 스펙을 만족할 가능성이 높은 입력 조건 (추천 작업 구간)</div>', unsafe_allow_html=True)
        xmin_d, xmax_d = float(np.floor(xa.min() * 100) / 100), float(np.ceil(xa.max() * 100) / 100)
        ymin_d, ymax_d = float(np.floor(ya.min() * 100) / 100), float(np.ceil(ya.max() * 100) / 100)
        p1, p2, p3, p4 = st.columns(4)
        target = p1.slider("목표 OK 확률 (%)", 50, 99, 90, key="cond_target")
        min_n = p2.number_input("최소 근거 시료 수", min_value=1, max_value=500, value=20, step=5, key="cond_minn",
                                help="조건 주변(아래 창 크기)에 실제로 측정된 시료가 이만큼 이상 있는 곳만 추천합니다 (데이터가 없는 곳을 추측으로 추천하지 않기 위해)")
        wx = p3.number_input(f"근거 창 {xk} ± (mm)", min_value=0.01, max_value=0.5, value=0.05, step=0.01, format="%.2f", key="cond_wx")
        wy = p4.number_input(f"근거 창 {yk} ± (mm)", min_value=0.01, max_value=0.5, value=0.05, step=0.01, format="%.2f", key="cond_wy")
        b1, b2, b3, b4 = st.columns(4)
        xlo_i = b1.number_input(f"{xk} 하한 (조정 가능 범위)", value=xmin_d, step=0.01, format="%.3f", key=f"cond_xlo_{xmode}")
        xhi_i = b2.number_input(f"{xk} 상한", value=xmax_d, step=0.01, format="%.3f", key=f"cond_xhi_{xmode}")
        ylo_i = b3.number_input(f"{yk} 하한 (조정 가능 범위)", value=ymin_d, step=0.01, format="%.3f", key="cond_ylo")
        yhi_i = b4.number_input(f"{yk} 상한", value=ymax_d, step=0.01, format="%.3f", key="cond_yhi")
        st.caption("조정 가능 범위의 기본값은 이번 데이터에서 실제 측정된 범위입니다. 현장에서 바꿀 수 있는 범위가 더 좁으면 줄여서 입력하세요.")

        if xlo_i >= xhi_i or ylo_i >= yhi_i:
            st.error("조정 가능 범위의 하한이 상한보다 작아야 합니다.")
        else:
            xlim_i, ylim_i = (float(xlo_i), float(xhi_i)), (float(ylo_i), float(yhi_i))
            cmap = condition_map(xa, ya, tm, lo, hi, quad, xlim_i, ylim_i, float(wx), float(wy), int(min_n), target / 100)
            rect = cmap["rect"]
            if rect is None:
                st.warning(f"OK 확률 {target}% 이상이면서 근거 시료 {int(min_n)}개 이상인 구간이 없습니다. 목표 확률이나 최소 근거 시료 수를 낮춰보세요.")
            else:
                q1, q2, q3, q4 = st.columns(4)
                q1.markdown(_kpi(f"추천 {xk} 구간", f"{rect['x'][0]:.2f} ~ {rect['x'][1]:.2f}", "mm"), unsafe_allow_html=True)
                q2.markdown(_kpi(f"추천 {yk} 구간", f"{rect['y'][0]:.2f} ~ {rect['y'][1]:.2f}", "mm"), unsafe_allow_html=True)
                q3.markdown(_kpi("구간 안 최소 OK 확률(모델)", f"{rect['p_min'] * 100:.1f}%", f"모든 칸이 목표 {target}% 이상"), unsafe_allow_html=True)
                real_rate = rect["ok_real"] / rect["n_real"] * 100 if rect["n_real"] else float("nan")
                q4.markdown(_kpi("구간 안 실측 시료", f"{rect['n_real']}개 중 {rect['ok_real']}개 OK", f"실측 OK 비율 {real_rate:.1f}%",
                                 color=SUCCESS if real_rate >= target else WARNING), unsafe_allow_html=True)
                st.markdown("<div style='margin-top:10px;'></div>", unsafe_allow_html=True)
                st.success(
                    f"{xk} {rect['x'][0]:.2f} ~ {rect['x'][1]:.2f} mm, {yk} {rect['y'][0]:.2f} ~ {rect['y'][1]:.2f} mm 로 맞추면 "
                    f"3개 스펙을 모두 만족할 확률이 모델 기준 {rect['p_min'] * 100:.0f}% 이상입니다. "
                    f"이번 데이터에서 이 구간에 들어온 실제 시료는 {rect['n_real']}개이고 그 중 {rect['ok_real']}개({real_rate:.1f}%)가 OK 였습니다.")
            show_pts = st.checkbox("실측 시료 점 표시", value=False, key="cond_show_pts")
            pts_df = pd.DataFrame({"x": xa, "y": ya, "status": np.where(np.all((tm >= lo) & (tm <= hi), axis=1), "OK", "NG")}) if show_pts else None
            st.altair_chart(pok_heatmap(cmap, int(min_n), rect, pts_df), use_container_width=True, theme=None)
            st.caption("칸의 색이 초록일수록 그 입력 조건에서 3개 스펙을 모두 만족할 확률(모델 추정)이 높습니다. 회색 칸은 주변에 실측 시료가 부족해서 "
                       "추천하지 않는 곳이고, 흰색 사각형이 추천 작업 구간입니다.")

            top_df = cmap["tops"]
            if len(top_df):
                for t, k in enumerate(tg):
                    top_df[f"예측 {k}"] = cmap["pred_top"][:, t]
                st.markdown('<div class="block-title" style="margin-top:12px;">추천 중심 조건 후보 (OK 확률 높은 순, 서로 떨어진 5곳)</div>', unsafe_allow_html=True)
                tcfg = {c: st.column_config.NumberColumn(format="%.2f") for c in (xk, yk)}
                tcfg.update({c: st.column_config.NumberColumn(format="%.1f") for c in ("OK 확률(모델, %)", "주변 실측 OK 비율(%)")})
                tcfg.update({f"예측 {k}": st.column_config.NumberColumn(format="%.3f") for k in tg})
                st.dataframe(top_df, use_container_width=True, hide_index=True, column_config=tcfg)
                st.caption("'주변 실측'은 그 조건 근처(근거 창 안)에서 실제로 측정된 시료의 수와 OK 비율입니다. 모델 확률과 실측 비율이 비슷할수록 믿을 만합니다. "
                           "이 데이터는 이미 만들어진 시료를 측정한 것이라, 추천 조건으로 시험 시료를 10~20개 만들어 실제 OK 비율을 확인하는 것을 권장합니다.")

            st.markdown("<hr>", unsafe_allow_html=True)

            # ---------------- ③ 현재 시료 조정 ----------------
            st.markdown('<div class="block-title">③ 현재 NG 시료의 입력 조건을 얼마나 바꾸면 스펙 안에 들어가는가 (추정)</div>', unsafe_allow_html=True)
            a1, a2, a3 = st.columns(3)
            limx = a1.number_input(f"{xk} 변경 한도 ± (mm)", min_value=0.0, max_value=3.0, value=1.0, step=0.05, format="%.2f", key="cond_limx",
                                   help="한 시료에서 바꿀 수 있는 최대 변경량입니다. 현장 제약에 맞게 줄이세요.")
            limy = a2.number_input(f"{yk} 변경 한도 ± (mm)", min_value=0.0, max_value=3.0, value=0.5, step=0.05, format="%.2f", key="cond_limy")
            kmargin = a3.slider("안전 여유 (예측 오차의 배수)", 0.0, 2.0, 0.5, 0.25, key="cond_k",
                                help="결과 예측값이 스펙 경계에서 (예측 오차 × 이 값) 만큼 안쪽에 들어오도록 요구합니다. 클수록 보수적입니다.")
            prio = st.radio("변경 우선순위", ["균형 (두 입력 모두 조금씩)", f"{xk} 위주로 변경", f"{yk} 위주로 변경"], horizontal=True, key="cond_prio")
            scales = (0.1, 0.05) if prio.startswith("균형") else ((0.1, 1.0) if prio.startswith(xk) else (1.0, 0.05))

            plan = plan_adjustments(xa, ya, tm, lo, hi, quad, cv["rmse"], xlim_i, ylim_i, float(limx), float(limy), float(kmargin), scales)
            n_ng = len(plan)
            if n_ng == 0:
                st.success("현재 NG 시료가 없습니다.")
            else:
                pos = plan["idx"].to_numpy()
                feas = plan["ok"].to_numpy()
                dx = plan["nx"].to_numpy() - xa[pos]
                dy = plan["ny"].to_numpy() - ya[pos]
                s1, s2, s3, s4 = st.columns(4)
                s1.markdown(_kpi("현재 NG 시료", f"{n_ng}개", "3개 스펙 중 하나 이상 벗어남"), unsafe_allow_html=True)
                s2.markdown(_kpi("조정으로 OK 가능(추정)", f"{int(feas.sum())}개", f"NG 의 {feas.mean() * 100:.0f}% · 불가 {int((~feas).sum())}개",
                                 color=SUCCESS if feas.mean() >= 0.7 else WARNING), unsafe_allow_html=True)
                s3.markdown(_kpi(f"{xk} 변경량 (중앙값)", f"{np.nanmedian(dx):+.3f} mm" if feas.any() else "-",
                                 f"사분위 {np.nanpercentile(dx, 25):+.2f} ~ {np.nanpercentile(dx, 75):+.2f}" if feas.any() else ""), unsafe_allow_html=True)
                s4.markdown(_kpi("변경 후 예상 OK 확률(평균)", f"{np.nanmean(plan['pok']):.1f}%" if feas.any() else "-", "그 조건에서의 일반적인 OK 확률(모델)"), unsafe_allow_html=True)

                view = pd.DataFrame({"시료": dfc["시료"].to_numpy()[pos], "엑셀 행": dfc["엑셀 행"].to_numpy()[pos],
                                     f"현재 {xk}": xa[pos], f"현재 {yk}": ya[pos]})
                for t, k in enumerate(tg):
                    view[f"현재 {k}"] = tm[pos, t]
                view["불합격 항목"] = [", ".join(k for t, k in enumerate(tg) if not (lo[t] <= tm[i, t] <= hi[t])) for i in pos]
                view[f"추천 {xk}"], view[f"추천 {yk}"] = plan["nx"].to_numpy(), plan["ny"].to_numpy()
                view[f"Δ{xk}"], view[f"Δ{yk}"] = dx, dy
                for t, k in enumerate(tg):
                    view[f"변경 후 예측 {k}"] = plan[f"p{t}"].to_numpy()
                view["예상 OK 확률(%)"] = plan["pok"].to_numpy()
                view["판정"] = np.where(feas, "조정 가능(추정)", "한도 내 조정 불가")
                vcfg = {c: st.column_config.NumberColumn(format="%.4f") for c in view.columns
                        if c.startswith(("현재 ", "추천 ", "변경 후")) and c != "현재 NG"}
                vcfg.update({f"Δ{xk}": st.column_config.NumberColumn(format="%+.3f"), f"Δ{yk}": st.column_config.NumberColumn(format="%+.3f"),
                             "예상 OK 확률(%)": st.column_config.NumberColumn(format="%.1f")})
                st.dataframe(view, use_container_width=True, hide_index=True, column_config=vcfg)
                st.caption("각 시료의 '변경 후 예측' 결과값은 그 시료만의 고유한 편차(다른 요인의 영향)가 그대로 유지된다는 가정으로 계산했습니다. "
                           "'예상 OK 확률'은 추천 조건에서 일반적인 시료가 스펙을 만족할 확률이라 더 보수적인 값입니다. "
                           "입력을 바꾸면 결과가 이 데이터의 경향대로 따라 바뀐다는 가정이 필요하므로, 실제 조정 후 재측정으로 확인하세요.")
                if feas.any():
                    cur_df = pd.DataFrame({"x": xa[pos][feas], "y": ya[pos][feas], "sample": dfc["시료"].to_numpy()[pos][feas]})
                    rec_df = pd.DataFrame({"x": plan["nx"].to_numpy()[feas], "y": plan["ny"].to_numpy()[feas], "sample": dfc["시료"].to_numpy()[pos][feas]})
                    st.altair_chart(adjust_arrow_chart(cur_df, rec_df), use_container_width=True, theme=None)
                    st.caption("빨간 점이 현재 NG 시료의 입력 조건, 초록 점이 추천 조건이며 회색 선이 이동 방향입니다.")

                rect_df = None
                if rect is not None:
                    rect_df = pd.DataFrame([{"전장 하한": rect["x"][0], "전장 상한": rect["x"][1], "일자-Y 하한": rect["y"][0], "일자-Y 상한": rect["y"][1],
                                             "구간 내 최소 OK 확률(모델,%)": rect["p_min"] * 100, "구간 안 실측 시료 수": rect["n_real"], "그 중 OK": rect["ok_real"]}])
                cls_df = pd.DataFrame({"": ["실제 OK", "실제 NG"], "예측 OK": [cv["tp"], cv["fp"]], "예측 NG": [cv["fn"], cv["tn"]]})
                params = {"예측 모델": mdl, "목표 OK 확률(%)": target, "최소 근거 시료 수": int(min_n), f"근거 창 {xk}±": wx, f"근거 창 {yk}±": wy,
                          f"조정 가능 범위 {xk}": f"{xlo_i} ~ {xhi_i}", f"조정 가능 범위 {yk}": f"{ylo_i} ~ {yhi_i}",
                          f"변경 한도 {xk}±": limx, f"변경 한도 {yk}±": limy, "안전 여유(오차 배수)": kmargin, "변경 우선순위": prio}
                st.download_button(
                    "조건찾기_결과.xlsx 다운로드",
                    data=build_cond_xlsx(acc_df, cls_df, cv["calib"], top_df, rect_df, view, params),
                    file_name="조건찾기_결과.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )


# ---------------- Tab 4. Go/No-Go 게이지 설계 ----------------
with tab_gauge:
    r = _analysis_result()
    if r is not None:
        trade = r.get("tradeoff")
        trade_end = r.get("tradeoff_end")
        gbasis = "c"
        if trade and trade_end:
            _gsel = st.radio("전장 기준", ["중심 기준 전장", "양 끝단 포함 전장 (지그 치수용)"], horizontal=True, key="gauge_basis",
                             help="양 끝단 포함 전장 = 오른쪽 끝단 최외측 Y − 왼쪽 끝단 최외측 Y (파이프 기울기로 튀어나온 길이 포함). 지그는 이 값을 기준으로 만들어야 제품이 들어갑니다.")
            if _gsel.startswith("양 끝단"):
                trade, gbasis = trade_end, "e"
        if not trade:
            st.info("합격 시료가 없거나 이전 버전의 결과입니다. Tab 1 에서 'Min/Max 분석 실행'을 다시 눌러주세요.")
        else:
            df_all = r["df"]
            flags_all = pass_matrix(df_all, r["specs"])["전체"].values
            xk, yk = trade[0]["xkey"], trade[0]["ykey"]

            st.markdown('<div class="block-title">① 허용 불합격률별 범위 비교</div>', unsafe_allow_html=True)
            st.caption(f"{xk} 범위와 {yk} 범위를 동시에 만족하는 시료 중 불합격이 차지하는 비율을 일정 수준까지 허용했을 때, "
                       "범위가 얼마나 넓어지고 게이지 통과율이 얼마나 올라가는지 비교합니다. "
                       "허용률 0% 는 불합격이 하나도 없는 범위(Tab 1 ④ 참고)입니다.")
            rows = []
            for w in trade:
                rows.append({
                    "허용 불합격률(%)": w["rate"] * 100,
                    f"{xk} 하한": w["x"][0], f"{xk} 상한": w["x"][1], f"{yk} 하한": w["y"][0], f"{yk} 상한": w["y"][1],
                    "범위 안 시료 수": w["n_in"], "그 중 불합격": w["n_fail_in"],
                    "실제 불합격률(%)": w["n_fail_in"] / w["n_in"] * 100 if w["n_in"] else 0.0,
                    "합격 시료 포함률(%)": w["n_pass_in"] / w["n_pass"] * 100,
                    "게이지 통과율(%)": w["n_in"] / w["n_total"] * 100,
                })
            tdf = pd.DataFrame(rows)
            cfg = {c: st.column_config.NumberColumn(format="%.4f") for c in tdf.columns if c.endswith(("하한", "상한"))}
            cfg.update({c: st.column_config.NumberColumn(format="%.1f") for c in ("허용 불합격률(%)", "실제 불합격률(%)", "합격 시료 포함률(%)", "게이지 통과율(%)")})
            st.dataframe(tdf, use_container_width=True, hide_index=True, column_config=cfg)

            long = pd.concat([
                pd.DataFrame({"rate": tdf["허용 불합격률(%)"].map(lambda v: f"{v:g}%"), "metric": "합격 시료 포함률", "v": tdf["합격 시료 포함률(%)"]}),
                pd.DataFrame({"rate": tdf["허용 불합격률(%)"].map(lambda v: f"{v:g}%"), "metric": "실제 불합격률", "v": tdf["실제 불합격률(%)"]}),
            ])
            tb = alt.Chart(long).mark_bar().encode(
                x=alt.X("rate:N", sort=None, title="허용 불합격률", axis=alt.Axis(labelAngle=0)),
                xOffset="metric:N", y=alt.Y("v:Q", title="%"),
                color=alt.Color("metric:N", title=None, scale=alt.Scale(domain=["합격 시료 포함률", "실제 불합격률"], range=[SUCCESS, DANGER])),
                tooltip=[alt.Tooltip("rate:N", title="허용률"), alt.Tooltip("metric:N", title="구분"), alt.Tooltip("v:Q", format=".1f", title="%")])
            st.altair_chart(_dark(tb, 260), use_container_width=True, theme=None)
            st.caption("허용률을 높이면 합격 시료를 더 많이 포함(양품 손실↓)하지만, 게이지를 통과한 부품 중 불합격이 섞이는 비율(불량 유출)이 늘어납니다.")

            labels = [f"허용 {w['rate'] * 100:g}% — {xk} {w['x'][0]:.3f}~{w['x'][1]:.3f} · {yk} {w['y'][0]:.3f}~{w['y'][1]:.3f} "
                      f"(시료 {w['n_in']}개, 불합격 {w['n_fail_in']}개)" for w in trade]
            pick = st.selectbox("게이지 설계에 사용할 범위", labels, index=0, key=f"gauge_pick_{gbasis}")
            win = trade[labels.index(pick)]

            st.markdown("<hr>", unsafe_allow_html=True)
            st.markdown('<div class="block-title">② 게이지 공차 반영 — GO/NO-GO 설계 치수</div>', unsafe_allow_html=True)
            st.caption("게이지 제작 공차 영역을 합격 범위 '안쪽'에 두어, 게이지가 공차 한계로 만들어져도 범위 밖 부품은 통과시키지 않도록 "
                       "설계 치수를 안쪽으로 줄입니다 (작업 공차 안쪽 배치 방식).")
            g1, g2 = st.columns(2)
            pct = g1.number_input("게이지 제작 공차 (범위 폭 대비 %)", min_value=0.0, max_value=45.0, value=10.0, step=1.0, key="gauge_pct")
            wear = g2.number_input("마모 여유 (범위 폭 대비 %)", min_value=0.0, max_value=45.0, value=0.0, step=1.0, key="gauge_wear")
            m1, m2 = st.columns(2)
            mm_x = m1.number_input(f"{xk} 제작 공차 직접 입력 (mm, 비우면 % 적용)", value=None, min_value=0.0, step=0.001, format="%.4f", key="gauge_mm_x")
            mm_y = m2.number_input(f"{yk} 제작 공차 직접 입력 (mm, 비우면 % 적용)", value=None, min_value=0.0, step=0.001, format="%.4f", key="gauge_mm_y")

            design = gauge_design(win, pct, wear, {xk: mm_x, yk: mm_y})
            dcfg = {c: st.column_config.NumberColumn(format="%.4f") for c in
                    ("합격 범위 하한", "합격 범위 상한", "범위 폭", "제작 공차(mm)", "마모 여유(mm)", "하한측 설계 치수", "상한측 설계 치수", "게이지 통과 폭")}
            st.dataframe(design, use_container_width=True, hide_index=True, column_config=dcfg)

            if not design["설계 가능"].all():
                st.error("제작 공차와 마모 여유의 합이 합격 범위 폭의 절반 이상이라 설계할 수 없습니다. 공차를 줄이거나 더 넓은 범위(허용률 상향)를 선택하세요.")
            else:
                xr = (float(design.loc[0, "하한측 설계 치수"]), float(design.loc[0, "상한측 설계 치수"]))
                yr = (float(design.loc[1, "하한측 설계 치수"]), float(design.loc[1, "상한측 설계 치수"]))
                chk = gauge_check(df_all, flags_all, xk, yk, xr, yr)
                loss_pct = chk["n_pass_lost"] / chk["n_pass"] * 100 if chk["n_pass"] else 0
                st.markdown('<div class="block-title">③ 이 게이지로 판정한 결과 (이번 시료 기준)</div>', unsafe_allow_html=True)
                c1, c2, c3, c4 = st.columns(4)
                c1.markdown(_kpi("게이지 통과(GO)", f"{chk['n_go']}개", f"전체의 {chk['n_go'] / chk['n_total'] * 100:.1f}%"), unsafe_allow_html=True)
                c2.markdown(_kpi("통과품 중 불합격(유출)", f"{chk['n_go_fail']}개",
                                 f"통과품의 {chk['n_go_fail'] / chk['n_go'] * 100 if chk['n_go'] else 0:.1f}%",
                                 color=SUCCESS if chk["n_go_fail"] == 0 else WARNING), unsafe_allow_html=True)
                c3.markdown(_kpi("양품인데 NG 판정", f"{chk['n_pass_lost']}개", f"양품의 {loss_pct:.1f}%"), unsafe_allow_html=True)
                c4.markdown(_kpi("게이지 통과 폭", f"{design.loc[0, '게이지 통과 폭']:.3f} / {design.loc[1, '게이지 통과 폭']:.3f}", f"{xk} / {yk} (mm)"), unsafe_allow_html=True)
                st.altair_chart(window_scatter_chart(df_all, flags_all, win, {"x": xr, "y": yr}), use_container_width=True, theme=None)
                st.caption("흰색 사각형은 선택한 합격 범위, 파란 점선 사각형은 게이지 공차를 반영한 설계 범위입니다. "
                           "측정 불확도(측정기 오차)와 온도 영향은 반영하지 않았으니 실제 제작 전에 별도로 검토하세요.")

                params = {"데이터 파일": r["file"], "선택한 범위": pick, "게이지 제작 공차(%)": pct, "마모 여유(%)": wear,
                          f"{xk} 제작 공차 직접 입력(mm)": mm_x, f"{yk} 제작 공차 직접 입력(mm)": mm_y}
                st.download_button(
                    "게이지_설계_결과.xlsx 다운로드",
                    data=build_gauge_xlsx(tdf, design, chk, params),
                    file_name="게이지_설계_결과.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )


# ---------------- Tab 5. 품질 분석 (아코디언 5개, 기본은 모두 닫힘) ----------------
with tab_analysis:
    r = _analysis_result()
    if r is not None:
        st.caption("항목을 눌러 펼치면 해당 분석이 표시됩니다 (기본은 모두 닫혀 있음). 전부 Tab 1 에서 실행한 같은 데이터로 계산합니다.")
        with st.expander("전장 구간별 합격률", expanded=False):
            df_all, specs = r["df"], r["specs"]
            st.markdown('<div class="block-title">분석 조건</div>', unsafe_allow_html=True)
            b1, b2, b3 = st.columns(3)
            width = b1.select_slider("전장 구간 폭 (mm)", options=[0.05, 0.1, 0.2, 0.25, 0.5], value=0.1, key="bin_width")
            target_pct = b2.slider("목표 합격률 (%)", 50, 100, 90, key="bin_target")
            min_n = b3.number_input("최소 시료 수 (신뢰 기준)", min_value=5, max_value=500, value=30, step=5, key="bin_min_n")

            table = length_bin_table(df_all, specs, width)
            curve = length_threshold_curve(df_all, specs)
            rec = recommend_threshold(curve, target_pct, int(min_n))
            overall = len(r["subset"]) / len(df_all) * 100 if len(df_all) else 0

            k1, k2, k3, k4 = st.columns(4)
            k1.markdown(_kpi("전체 합격률", f"{overall:.1f}%", f"{len(r['subset'])}개 / {len(df_all)}개"), unsafe_allow_html=True)
            if rec is not None:
                k2.markdown(_kpi("권장 전장 하한", f"{rec['전장 하한']:.2f} mm", f"합격률 {target_pct}% 이상 · 시료 {int(min_n)}개 이상", color=SUCCESS), unsafe_allow_html=True)
                k3.markdown(_kpi("그 하한에서의 합격률", f"{rec['합격률(%)']:.1f}%", f"전체 대비 {rec['합격률(%)'] - overall:+.1f}%p"), unsafe_allow_html=True)
                k4.markdown(_kpi("그 하한 이상 시료 수", f"{int(rec['시료 수'])}개", f"전체의 {rec['시료 수'] / len(df_all) * 100:.0f}%"), unsafe_allow_html=True)
            else:
                k2.markdown(_kpi("권장 전장 하한", "없음", "목표를 만족하는 구간 없음", color=WARNING), unsafe_allow_html=True)
                ok_curve = curve[curve["시료 수"] >= int(min_n)]
                best = ok_curve.loc[ok_curve["합격률(%)"].idxmax()] if not ok_curve.empty else None
                k3.markdown(_kpi("달성 가능한 최대 합격률", f"{best['합격률(%)']:.1f}%" if best is not None else "-",
                                 f"전장 {best['전장 하한']:.2f} mm 이상" if best is not None else ""), unsafe_allow_html=True)
                k4.markdown(_kpi("목표 합격률", f"{target_pct}%"), unsafe_allow_html=True)

            st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
            st.markdown('<div class="block-title">① 전장 구간별 합격률</div>', unsafe_allow_html=True)
            cdf = pd.DataFrame({"label": table["전장 구간"], "rate": table["합격률(%)"], "n": table["시료 수"], "ok": table["합격 수"]})
            cdf["level"] = np.where(cdf["n"] < int(min_n), "표본 부족",
                                    np.where(cdf["rate"] >= target_pct, "목표 이상", "목표 미만"))
            bars = alt.Chart(cdf).mark_bar().encode(
                x=alt.X("label:N", sort=None, title="전장 구간 (mm)"),
                y=alt.Y("rate:Q", title="합격률 (%)", scale=alt.Scale(domain=[0, 105])),
                color=alt.Color("level:N", title="구분", scale=alt.Scale(domain=["목표 이상", "목표 미만", "표본 부족"], range=[SUCCESS, ACCENT_BLUE, "#4b5563"])),
                tooltip=[alt.Tooltip("label:N", title="전장 구간"), alt.Tooltip("rate:Q", format=".1f", title="합격률(%)"),
                         alt.Tooltip("ok:Q", title="합격 수"), alt.Tooltip("n:Q", title="시료 수")])
            labels = alt.Chart(cdf).mark_text(dy=-8, color="#cbd5e1", fontSize=11).encode(
                x=alt.X("label:N", sort=None), y="rate:Q", text=alt.Text("n:Q"))
            goal = alt.Chart(pd.DataFrame({"y": [target_pct], "kind": ["목표"]})).mark_rule(strokeDash=[6, 4], strokeWidth=1.8).encode(
                y="y:Q", color=alt.Color("kind:N", scale=_line_color_scale(), legend=None))
            st.altair_chart(_dark(alt.layer(bars, labels, goal), 340), use_container_width=True, theme=None)
            st.caption("막대 위 숫자는 구간의 시료 수입니다. 회색 막대는 시료 수가 최소 기준보다 적어 신뢰하기 어려운 구간입니다.")

            st.markdown('<div class="block-title">② 전장 하한을 올렸을 때의 합격률 (누적)</div>', unsafe_allow_html=True)
            cc = curve[curve["시료 수"] >= int(min_n)].rename(columns={"전장 하한": "thr", "시료 수": "n", "합격률(%)": "rate"})
            if cc.empty:
                st.info("최소 시료 수 기준을 만족하는 구간이 없습니다. 기준을 낮춰보세요.")
            else:
                cl = alt.Chart(cc).mark_line(color=ACCENT_BLUE, strokeWidth=2.5, point=alt.OverlayMarkDef(color=ACCENT_BLUE, size=30)).encode(
                    x=alt.X("thr:Q", title="전장 하한 (mm) — 이 값 이상인 시료만 남김", scale=alt.Scale(zero=False)),
                    y=alt.Y("rate:Q", title="합격률 (%)", scale=alt.Scale(domain=[0, 105])),
                    tooltip=[alt.Tooltip("thr:Q", format=".2f", title="전장 하한"), alt.Tooltip("rate:Q", format=".1f", title="합격률(%)"),
                             alt.Tooltip("n:Q", title="남은 시료 수")])
                layers = [cl, alt.Chart(pd.DataFrame({"y": [target_pct], "kind": ["목표"]})).mark_rule(strokeDash=[6, 4], strokeWidth=1.8).encode(
                    y="y:Q", color=alt.Color("kind:N", scale=_line_color_scale(), legend=None))]
                if rec is not None:
                    layers.append(alt.Chart(pd.DataFrame({"x": [rec["전장 하한"]], "kind": ["기준값"]})).mark_rule(strokeWidth=2).encode(
                        x="x:Q", color=alt.Color("kind:N", scale=_line_color_scale(), legend=None)))
                st.altair_chart(_dark(alt.layer(*layers), 320), use_container_width=True, theme=None)

            st.markdown('<div class="block-title">③ 구간별 상세 표</div>', unsafe_allow_html=True)
            show = table.drop(columns=["구간 시작", "구간 끝"])
            cfg = {c: st.column_config.NumberColumn(format="%.1f") for c in show.columns if c.endswith("(%)")}
            st.dataframe(show, use_container_width=True, hide_index=True, column_config=cfg)
            st.caption("권장 전장 하한은 '이 값 이상으로 관리하면 합격률이 목표에 도달한다'는 통계적 경향입니다. "
                       "전장은 규격(공차)이 없는 관리 변수이므로, 다른 품질 항목과의 균형은 공정 담당자가 함께 판단하세요.")
        with st.expander("공정능력 (Cpk)", expanded=False):
            data, choice = _pick_dataset(r, "cap_dataset")
            if choice == "스펙 합격 시료":
                st.warning("합격 시료만으로 계산하면 규격 밖 시료가 이미 제외돼 있어 능력이 실제보다 좋게 나옵니다. "
                           "공정능력은 '전체 유효 시료' 기준을 우선 보세요.")
            st.caption("규격(상한)이 있는 타겟 항목(단축·장축 위치도, 단축 윤곽도)만 계산합니다. "
                       "전장·일자-Y 는 규격이 없는 관리 변수라서 공정능력 대상이 아닙니다 (분포·상관·관리도 탭에서 분석).")

            rows = []
            for k, (lo, hi) in r["specs"].items():
                lo, hi = (lo, hi) if lo <= hi else (hi, lo)
                lsl = lo if lo > 0 else None  # 0 은 물리적 하한이므로 규격 하한으로 보지 않는다 (한쪽 규격)
                c = capability(data[k], lsl, hi)
                rows.append({
                    "항목": k, "규격(상한)": f"{hi:g}" if lsl is None else f"{lsl:g} ~ {hi:g}", "시료 수": c["n"],
                    "평균": c["mean"], "표준편차": c["std"], "Cpk(정규, Cpu)": c["cpk"], "Cpk(백분위)": c["cpk_pct"],
                    "판정(Cpk 기준)": cap_rating(c["cpk"]), "예상 불량 PPM(정규)": c["ppm"], "실측 규격 밖": c["n_out"],
                })
            cap_df = pd.DataFrame(rows)

            st.markdown('<div class="block-title">① 타겟 항목별 공정능력</div>', unsafe_allow_html=True)
            cfg = {c: st.column_config.NumberColumn(format="%.3f") for c in ("평균", "표준편차")}
            cfg.update({c: st.column_config.NumberColumn(format="%.2f") for c in ("Cpk(정규, Cpu)", "Cpk(백분위)")})
            cfg["예상 불량 PPM(정규)"] = st.column_config.NumberColumn(format="%.0f")
            st.dataframe(cap_df, use_container_width=True, hide_index=True, column_config=cfg)

            plot = cap_df.dropna(subset=["Cpk(정규, Cpu)"])
            if not plot.empty:
                st.markdown('<div class="block-title">② Cpk 비교 (기준선: 1.00 / 1.33 / 1.67)</div>', unsafe_allow_html=True)
                cb = alt.Chart(plot).mark_bar(color=ACCENT_BLUE).encode(
                    x=alt.X("항목:N", sort=None, title=None, axis=alt.Axis(labelAngle=0)), y=alt.Y("Cpk(정규, Cpu):Q", title="Cpk"),
                    tooltip=[alt.Tooltip("항목:N"), alt.Tooltip("Cpk(정규, Cpu):Q", format=".2f", title="Cpk"),
                             alt.Tooltip("판정(Cpk 기준):N", title="판정")])
                ref = pd.DataFrame({"y": [1.0, 1.33, 1.67], "kind": ["LSL", "기준값", "목표"]})
                rl = alt.Chart(ref).mark_rule(strokeDash=[6, 4], strokeWidth=1.6).encode(
                    y="y:Q", color=alt.Color("kind:N", scale=_line_color_scale(), legend=None))
                st.altair_chart(_dark(alt.layer(cb, rl), 300), use_container_width=True, theme=None)

            st.caption(
                "위치도·윤곽도는 0 이하로 내려갈 수 없는 한쪽(상한) 규격이라 Cp 는 없고 Cpk = Cpu = (USL−μ) / 3σ 만 계산합니다 (σ = 전체 표본 표준편차). "
                "분포가 0 쪽으로 몰려 한쪽으로 치우쳐 있어, 정규분포 가정 Cpk 와 백분위 방식 Cpk(중앙값에서 99.865% 점까지 거리 기준)를 함께 보여줍니다. "
                "예상 불량 PPM 은 정규분포 가정 값이라 치우친 분포에서는 실측 규격 밖 수와 차이가 날 수 있습니다.")
        with st.expander("분포 분석", expanded=False):
            data, choice = _pick_dataset(r, "dist_dataset")
            items = _analysis_items(r)
            st.markdown('<div class="block-title">① 기술 통계 · 정규성</div>', unsafe_allow_html=True)
            stat_df = pd.DataFrame([{"항목": k, **describe_item(data[k])} for k in items])
            stat_df["정규성 판정"] = np.where(stat_df["정규성 p값(JB)"] < 0.05, "정규분포와 유의하게 다름", "정규성 기각 안 됨")
            cfg = {c: st.column_config.NumberColumn(format="%.4f") for c in ("평균", "표준편차", "최소", "중앙값", "최대")}
            cfg.update({"왜도": st.column_config.NumberColumn(format="%.2f"), "첨도(초과)": st.column_config.NumberColumn(format="%.2f"),
                        "정규성 p값(JB)": st.column_config.NumberColumn(format="%.4f")})
            st.dataframe(stat_df, use_container_width=True, hide_index=True, column_config=cfg)
            st.caption("왜도 > 0 이면 오른쪽(큰 값 쪽)으로 꼬리가 깁니다. 시료가 수백 개 이상이면 정규성 검정(Jarque-Bera)은 작은 차이에도 "
                       "기각되므로, 아래 히스토그램 모양을 함께 보고 판단하세요.")

            st.markdown('<div class="block-title">② 히스토그램 · 박스플롯</div>', unsafe_allow_html=True)
            sel = st.selectbox("그래프 항목", items, key="dist_item")
            lines = _spec_lines_for(r, sel)
            if len(data) >= 3:
                st.altair_chart(hist_chart(data[sel], sel, lines), use_container_width=True, theme=None)
                st.altair_chart(box_compare_chart(r["df"][sel], r["subset"][sel], sel), use_container_width=True, theme=None)
                st.caption(f"히스토그램은 '{choice}' 기준이며, 파란 선은 같은 평균·표준편차의 정규분포입니다. "
                           "박스플롯은 전체와 합격 시료를 비교합니다 (상자: 25~75%, 가운데 선: 중앙값).")
                if sel in r["var_keys"]:
                    st.caption("전장·일자-Y 는 규격(공차)이 없는 관리 변수라 초록 점선은 엑셀 머리글의 기준값(참고용)입니다.")
            else:
                st.info("시료가 너무 적어 그래프를 그릴 수 없습니다.")
        with st.expander("상관 · 산점도", expanded=False):
            df_all, items = r["df"], _analysis_items(r)
            st.markdown('<div class="block-title">① 상관 행렬 (전체 유효 시료, Pearson r)</div>', unsafe_allow_html=True)
            st.altair_chart(corr_heatmap(df_all, items), use_container_width=True, theme=None)
            cm = df_all[items].corr()
            pairs = [{"항목 A": a, "항목 B": b, "상관계수 r": cm.loc[a, b]} for i, a in enumerate(items) for b in items[i + 1:]]
            pairs_df = pd.DataFrame(pairs).sort_values("상관계수 r", key=lambda s: s.abs(), ascending=False)
            pairs_df["강도"] = pd.cut(pairs_df["상관계수 r"].abs(), [-0.01, 0.3, 0.5, 0.7, 1.01], labels=["약함", "보통", "강함", "매우 강함"]).astype(str)
            st.dataframe(pairs_df, use_container_width=True, hide_index=True,
                         column_config={"상관계수 r": st.column_config.NumberColumn(format="%.2f")})
            st.caption("|r| 0.3 미만 약함, 0.5 이상 강함, 0.7 이상 매우 강함. 상관은 인과가 아니므로, 강한 쌍은 원인 후보로 보고 공정에서 확인하세요.")

            st.markdown('<div class="block-title">② 산점도 + 회귀선</div>', unsafe_allow_html=True)
            sx, sy = st.columns(2)
            x_item = sx.selectbox("X 축", items, index=0, key="sc_x")
            y_default = items.index("단축 윤곽도") if "단축 윤곽도" in items else min(1, len(items) - 1)
            y_item = sy.selectbox("Y 축", items, index=y_default, key="sc_y")
            flags = pass_matrix(df_all, r["specs"])["전체"].values
            sdf = pd.DataFrame({"x": df_all[x_item].values, "y": df_all[y_item].values, "sample": df_all["시료"].values,
                                "status": np.where(flags, "합격", "불합격")})
            if x_item == y_item:
                st.info("X 축과 Y 축에 서로 다른 항목을 선택해주세요.")
            else:
                slope, intercept = np.polyfit(sdf["x"], sdf["y"], 1)
                rr = float(np.corrcoef(sdf["x"], sdf["y"])[0, 1])
                m1, m2, m3 = st.columns(3)
                m1.markdown(_kpi("상관계수 r", f"{rr:.3f}"), unsafe_allow_html=True)
                m2.markdown(_kpi("결정계수 R²", f"{rr ** 2:.3f}", "X 로 설명되는 Y 변동 비율"), unsafe_allow_html=True)
                m3.markdown(_kpi("회귀 기울기", f"{slope:.4f}", f"{x_item} 1 증가 시 {y_item} {slope:+.4f}"), unsafe_allow_html=True)
                st.altair_chart(scatter_chart(sdf, x_item, y_item, slope, intercept), use_container_width=True, theme=None)
                st.caption("점 색은 3개 타겟 스펙을 모두 만족했는지(합격/불합격)입니다. 흰 점선은 단순 선형 회귀선입니다.")
        with st.expander("런차트 · 관리도", expanded=False):
            df_all, items = r["df"], _analysis_items(r)
            i1, i2 = st.columns([2, 1])
            sel = i1.selectbox("항목", items, key="imr_item")
            window = i2.slider("이동평균 구간 (시료 수)", 10, 100, 30, step=5, key="imr_window")
            st.caption("엑셀 행 순서를 시간(생산/측정) 순서로 보고 그립니다. 행 순서가 시간 순서가 아니면 해석에 주의하세요. "
                       "전체 유효 시료 기준입니다.")
            vals = df_all[sel].values
            if len(vals) < 10:
                st.info("시료가 너무 적어 관리도를 그릴 수 없습니다.")
            else:
                info_ = imr_stats(vals)
                n1, n2 = int(info_["rule1"].sum()), int(info_["rule2"].sum())
                trend_r = float(np.corrcoef(np.arange(len(vals)), vals)[0, 1])
                q1, q2, q3, q4 = st.columns(4)
                q1.markdown(_kpi("중심선 (평균)", f"{info_['mean']:.4f}", f"관리한계 {info_['lcl']:.4f} ~ {info_['ucl']:.4f}"), unsafe_allow_html=True)
                q2.markdown(_kpi("관리한계 밖 점", f"{n1}개", "규칙 1: ±3σ(I-MR) 초과", color=DANGER if n1 else SUCCESS), unsafe_allow_html=True)
                q3.markdown(_kpi("연속 8점 한쪽 치우침", f"{n2}개", "규칙 2: 평균 한쪽에 8점 이상", color=DANGER if n2 else SUCCESS), unsafe_allow_html=True)
                q4.markdown(_kpi("순서-값 상관", f"{trend_r:+.2f}", "|r| 클수록 경향(드리프트) 의심"), unsafe_allow_html=True)

                st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
                st.markdown('<div class="block-title">개별값(I) 관리도 · 런차트</div>', unsafe_allow_html=True)
                ic, mc = imr_charts(df_all["엑셀 행"].values, vals, sel, info_, _spec_lines_for(r, sel), window)
                st.altair_chart(ic, use_container_width=True, theme=None)
                st.markdown('<div class="block-title">이동범위(MR) 관리도</div>', unsafe_allow_html=True)
                st.altair_chart(mc, use_container_width=True, theme=None)
                st.caption("관리한계 = 평균 ± 2.66 × 평균이동범위(MR̄), MR 상한 = 3.267 × MR̄ (I-MR 표준 계수). 빨간 점은 규칙 1 또는 2 위반입니다. "
                           "파란 굵은 선은 이동평균이며, 한쪽으로 계속 움직이면 공정 평균이 이동(드리프트)하고 있다는 신호입니다. "
                           "관리한계는 '공정이 낼 수 있는 산포'이고 규격과는 다른 개념입니다.")


if tab_history is not None:
  with tab_history:
    st.markdown('<div class="block-title">① 처리 이력</div>', unsafe_allow_html=True)
    st.caption("이번 접속 세션 동안 'Min/Max 분석 실행'을 누른 기록입니다 (브라우저를 새로고침하거나 앱이 재시작되면 초기화됩니다).")
    history = st.session_state.get("process_history", [])
    if not history:
        st.info("아직 처리 이력이 없습니다. 왼쪽 사이드바에서 파일을 업로드하고 실행해보세요.")
    else:
        h1, h2, h3 = st.columns(3)
        h1.markdown(_kpi("이번 세션 분석 횟수", f"{len(history)}회"), unsafe_allow_html=True)
        h2.markdown(_kpi("최근 분석 시료 수", f"{history[0]['시료 수']}개"), unsafe_allow_html=True)
        h3.markdown(_kpi("최근 분석 파일", history[0]["데이터 파일"]), unsafe_allow_html=True)
        st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
        st.dataframe(pd.DataFrame(history), use_container_width=True, hide_index=True)
        if st.button("이력 초기화"):
            st.session_state["process_history"] = []
            st.rerun()

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<div class="block-title">② 자주 쓰는 분석 조건 (누적 통계)</div>', unsafe_allow_html=True)
    counter = st.session_state.get("spec_counter", Counter())
    if not counter:
        st.info("아직 집계된 통계가 없습니다.")
    else:
        stat_df = pd.DataFrame(counter.most_common(20), columns=["분석 조건", "분석 횟수"])
        c_left, c_right = st.columns([1, 1])
        with c_left:
            st.dataframe(stat_df, use_container_width=True, hide_index=True)
        with c_right:
            st.bar_chart(stat_df.set_index("분석 조건"))
        if st.button("통계 초기화"):
            st.session_state["spec_counter"] = Counter()
            st.rerun()

if tab_sync is not None:
  with tab_sync:
    st.markdown('<div class="block-title">분석 설정 관리 (오너 전용)</div>', unsafe_allow_html=True)
    st.caption(
        "타겟 항목은 엑셀 1행(구분 행)의 라벨, 변수 항목은 2행 머리글 텍스트에서 열을 찾는 정규식입니다. "
        "시료 행으로 인정할 파일명 패턴도 여기서 바꿉니다. 표기가 바뀌어 열을 못 찾을 때 정규식만 고치면 됩니다 (코드 수정 불필요)."
    )
    new_var = {k: st.text_input(f"변수 '{k}' 머리글 정규식", value=p, key=f"vrule_{k}")
               for k, p in st.session_state["var_rules"].items()}
    new_tgt = {k: st.text_input(f"타겟 '{k}' 라벨 정규식", value=p, key=f"trule_{k}")
               for k, p in st.session_state["target_rules"].items()}
    new_end = {k: st.text_input(f"{k} 원주 8점 머리글 정규식 (끝에 번호 괄호 하나)", value=p, key=f"erule_{k}")
               for k, p in st.session_state["end_rules"].items()}
    new_name_pat = st.text_input("시료 행 파일명 패턴 (정규식)", value=st.session_state["name_pattern"], key="rule_name_pattern")
    if st.button("분석 설정 저장"):
        try:
            for p in list(new_var.values()) + list(new_tgt.values()) + list(new_end.values()) + [new_name_pat]:
                _re.compile(p)
            st.session_state["var_rules"] = {k: v.strip() for k, v in new_var.items()}
            st.session_state["target_rules"] = {k: v.strip() for k, v in new_tgt.items()}
            st.session_state["end_rules"] = {k: v.strip() for k, v in new_end.items()}
            st.session_state["name_pattern"] = new_name_pat.strip()
            st.success("저장되었습니다. 다음 분석부터 적용됩니다.")
        except _re.error as e:
            st.error(f"정규식 오류: {e}")

    if data_bytes is not None:
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<div class="block-title">현재 업로드 파일에서 열 인식 테스트</div>', unsafe_allow_html=True)
        try:
            _raw = _load_raw_sheet(data_bytes, sheet_name)
            _hr, _cols = locate_columns(_raw, st.session_state["var_rules"], st.session_state["target_rules"])
            st.dataframe(_cols_table(_cols), use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(str(e))

if tab_help is not None:
  with tab_help:
    st.markdown("""
- **계산 기준**: 타겟 항목 3개(단축 위치도 0~1.0, 장축 위치도 0~1.0, 단축 윤곽도 0~0.8, 경계 포함)를 **모두 동시에** 만족하면 합격입니다.
  타겟 스펙 값은 Tab 1 상단에서 바꿀 수 있습니다.
- **최종 Min/Max (Tab 1 ①)**: 시료를 하나씩 보고 3개 타겟 스펙을 모두 만족하면 OK 로 분류한 뒤, OK 시료 전체에서 전장·일자-Y 의 가장 작은 값과
  가장 큰 값을 구합니다. OK 시료는 전부 이 범위 안에 들어옵니다. 반대로 이 범위 안에 들어온 시료가 전부 OK 인 것은 아닙니다.
- **양 끝단 포함 전장 (Tab 1 ②)**: 전장은 양쪽 끝단 '중심' 사이의 거리(휨 X − 일자 X)이고, 끝단 파이프는 기울어져 있어 원주 8점의 Y 값이 서로 다릅니다.
  그래서 지그를 만들 때는 전장에 기울어서 튀어나온 길이를 더해야 제품이 들어갑니다. 합격(OK) 시료마다 (오른쪽 끝단 Y-값_장1~8 중 가장 큰 값) − (왼쪽 끝단 Y-값_단1~8 중 가장 작은 값)을
  '양 끝단 포함 전장'으로 계산하고, 그 최소·최대를 보여줍니다. '안쪽 전장'(양 끝단이 가장 가까워지는 길이)과 기울기로 늘어난 길이도 함께 나옵니다.
- **합격 보장 범위 (Tab 1 ④ 참고, Tab 4 게이지 설계)**: 전장 범위와 일자-Y 범위를 **동시에** 만족하면 3개 타겟 스펙을 모두 만족하도록, 불합격 시료가 하나도 포함되지
  않는 직사각형 범위 중 합격 시료를 가장 많이 포함하는 범위를 정확히 찾습니다. 이번 데이터에서 실제 측정된 시료 기준이라 시료가 늘면 범위가 바뀔 수 있고,
  측정하지 않은 다른 요인의 영향은 반영되지 않습니다.
- **Tab 2 (조건 찾기)**: 전장·일자-Y 를 입력, 위치도 2종·윤곽도를 결과로 보고 이번 데이터의 경향으로 '3개 스펙을 만족하려면 입력을 어떻게 맞춰야 하는지'를 찾습니다.
  ① 예측 정확도(교차검증 오차 RMSE, 설명력 R², OK/NG 판정 정확도, 확률 보정), ② 3개 스펙을 만족할 확률이 목표(기본 90%) 이상이면서 실측 근거 시료가 충분한 입력 구간(추천 작업 구간)과
  OK 확률 지도, ③ 현재 NG 시료의 입력을 얼마나 바꾸면 스펙 안에 들어가는지(변경 한도·안전 여유·우선순위 설정 가능)를 보여줍니다.
  이 데이터는 이미 만들어진 시료를 측정한 관측 데이터라서 원인-결과가 확정된 것이 아니므로, 추천 조건은 시험 시료로 확인해야 합니다.
- **전장 기준 선택 (Tab 2, Tab 4)**: 두 탭 맨 위의 '전장 기준'에서 '양 끝단 포함 전장'을 고르면 전장 대신 양 끝단(파이프 기울기)을 포함한 전장으로 조건 찾기와 게이지 설계를 합니다. 지그 치수는 이 기준으로 만들어야 합니다.
- **Tab 3 (시료 확인)**: 스펙을 모두 만족한 시료 전체 표(전장 오름차순)와 엑셀 머리글에서 인식한 열 정보를 함께 보여줍니다.
- **Tab 4 (Go/No-Go 게이지 설계)**: ① 허용 불합격률(0~5%)별로 전장·일자-Y 범위를 비교하고, ② 게이지 제작 공차와 마모 여유를 안쪽으로 반영해 GO/NO-GO 설계 치수를 계산하고,
  ③ 그 게이지로 이번 시료를 판정한 결과(통과·유출·양품 NG)를 보여줍니다. 측정 불확도와 온도 영향은 반영하지 않습니다.
- **Tab 5 (품질 분석)**: 5개 아코디언(기본은 닫힘) — 전장 구간별 합격률, 공정능력(Cpk, 규격이 있는 타겟 3개만), 분포 분석, 상관·산점도, 런차트·관리도.
  주의: 이 엑셀은 행이 단축 윤곽도 오름차순으로 정렬되어 있어서, 런차트·관리도의 '엑셀 행 순서' 경향은 생산 순서가 아니라 정렬 때문일 수 있습니다.
- **제외되는 행**: 엑셀 하단의 `MIN / MAX / AVG`, `0.8 미만`, `0.8 이상`, `합격 / 불합격` 같은 요약·구분 행은 시료가 아니므로
  파일명 패턴(기본 `.xls` 또는 `QMS` 포함)으로 걸러냅니다.
- **열 찾기**: 타겟은 1행(구분 행) 라벨 위치, 변수는 2행 머리글 텍스트로 찾습니다. 열 위치가 바뀌어도 동작하지만
  표기가 바뀌면 '분석 설정 관리' 탭에서 정규식을 고쳐야 합니다.
- **한계**: 한 번에 시트 하나만 분석합니다. 시료 하나에 값이 하나라도 비어 있으면 그 행은 제외합니다.
- '처리 이력 · 통계' 탭은 이번 접속 세션 동안의 기록만 보여줍니다.
""")
