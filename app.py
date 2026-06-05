import os, sys, io, base64, threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl

CODE_CJ = 4
CODE_KD = 22

# 템플릿 파일 (base64 내장)
TMPL_B64 = "UEsDBBQAAAAIABE6tlxGx01IlQAAAM0AAAAQAAAAZG9jUHJvcHMvYXBwLnhtbE3PTQvCMAwG4L9SdreZih6kDkQ9ip68zy51hbYpbYT67+0EP255ecgboi6JIia2mEXxLuRtMzLHDUDWI/o+y8qhiqHke64x3YGMsRoPpB8eA8OibdeAhTEMOMzit7Dp1C5GZ3XPlkJ3sjpRJsPiWDQ6sScfq9wcChDneiU+ixNLOZcrBf+LU8sVU57mym/8ZAW/B7oXUEsDBBQAAAAIABE6tlyiCvEHMQEAAHsCAAARAAAAZG9jUHJvcHMvY29yZS54bWzFkstOwzAQRX+lyj61E/chWaklHkUsqIigEo+d5UxbizixbEPaLWLLvnwj6j+QmDSlgj3LuXPnzB1pEqGpKA2kptRgnATbW6u8sFToSbByTlOErFiB4rZfO4q6uSiN4q4uzRJpLp74ElCM8QgpcDzjjqMGGOqOGLAkE1QY4K40LT4THV4/m9zDMoEgBwWFsyjqRyhgu4/3z9ftbvuWoAOhoTkwyn4LkHVIr/7J9R0UtM61lZ2rqqp+RbyvPiJC97OrW39vKAvreCGgnrKSuo2GSbDffEfOzucXAYtxTEI8DqNojod0QGhMHpusR/kOgVWZyYX878SjEA/DOJ7jMY1GlAx+JN4HZEn9Fzm3btYKpxuWnqTTm/Q67V1OHxL0u9+MGHiRVpYFI97Rlb46/jP2BVBLAwQUAAAACAAROrZcda6PvKQGAAADIQAAEwAAAHhsL3RoZW1lL3RoZW1lMS54bWztWVuLGzcUfi/0P4h5d+buS4g32GO7abKbhOwmpY/asexRrBkZSd6NCYGSUGihFApp6Uuhb3kopYUGWvrSH7PQ0Kb9Dz0zvo1sOdk0m9LStcEzkr5z9Omco6PjmUuX76UMHREhKc+alnvBsRDJYt6n2bBp3T7oVeoWkgpnfcx4RprWlEjr8s7bb13CF1VCUoJAPpMXcdNKlBpftG0ZQzeWF/iYZDA24CLFCppiaPcFPga9KbM9x6naKaaZhTKcgtobgwGNCfIc10cVuHge+vPDT54/+cjaWczUZfCTKZl3xEzsx8X028QLuf7IzS9yKiMm0BFmTQsI9PnxAbmnLMSwVDDQtJziY9k7l+ylEFNbZEtyveIzl5sL9EdeISeGh0vBIAiDamup35vp38R1a91qt7rUVwBwHMOqXYPOmhcFc2wJNLs16O7UOr6r4Uv6/Q18K8y/Gt5f4YMNfK8XrWxYAs1uww182G60O7r+cIWvbuBrTqsT1DR8AUoYzUYbaCes+tFitUvIgLMrRngjDHo1bw5foexSpM3kM3WauEvxXS56AC4cjRXNkJqOyQDHIBNhRg8FRbt0mEAQjnHGJXQ7ntNzfPjNv0FxV3gXXyS4JD3riuVGV84NyVjQsWpaV0GrVYL8+tNPJw+fnjz88eTRo5OH383n3pS7grNhWe75k8/++OoD9PsPXz9//LkZL8v4Z99+/OznX16kXmm0vvj+2dPvf/3y09++eWyAtwQ+LMMPaEokuk6O0S2ewgINE5BD8WoSBwmmmgROAGkAdlWiAa9PMTPh2kQ34R0BWcMEfGdyV+O6n4iJogbgtSTVgHucszYXxuVcy+cqL2eSDc2Ti0kZdwvjI9Pc0ZqDu5MxhD81qYwSotG8ycDbeEgyolA+xkeEGMTep1Sz6x6NBZd8oND7FLUxNZrkgB4qs9AVmoJfpiaC4GrNNnt3UJszk/oOOdKRsC0wM6kkTDPjO3iicGpkjFNWRu5ilZhI7k9FrBlcKvD0kDCOun0ipUnmhphqdK9BhjG7fY9NUx0pFB2ZkLuY8zKyw0dRgtOxkTPNkjL2XTmCEMXoJldGElzfIXkb/ICzre6+Q4l6tW19GzKQOUDykYkwbQnC9f04ZQNMTMpbItWya0tQY3S0J0MttHcJYfgY9wlBt9814fmYm0lfTSCrXCEm21zFeqzm7YxIgooax+BYKrWQ3SdDvoXP3nQt8UxxlmKxTfP1kR4yXTjljKn0BotHWiqlIt+0ZhI3ZIpPpfVmgrWwytvSHK9Tkb3qHgOZu39DhryyDCT2U9vmADNiDpgDDAWGKd2CyMQskm+nQmxilBvom3blBnut3klp9tLiZ63sCf+ZsueNFTxnX+psSynrBc423H+wrOngSXaTwElyXtWcVzX/x6pm214+r2XOa5nzWuYfq2VW5YtdfuJTaElP9fhnQBnbV1NGdmVRBEnIA/0edBaNQsHyydM4gdv51BpuKHBxjwRX71GV7Cd4DFO6xQxDOVc9lGjMJZRR1lbdRRk2Sfd4f9bruouHnSCA1aofyrBFPxRtatZbra2e6i3VF62hLBMIC6WnJ1GaTCfhG0jU/NORcJ2zYtEwsKi7L2Jhl7wCBxXC+QP0MJgxgtCD8O7nfprJL7x75p7eZkx92Z5heY3gzDytkSiFm06iFIYJHCTr3Wfs60bD7GrPSKNWfxO+tjdzA8v0FjqGPeeHoCbG46Y1gD9QcJuOQZ/MsxZmw6xpxWpu6L+TWcZCqg6WyQxWDM3Wn1JFBGI0hVgvu4FlK26uV3P+veQazr/Pcva6k8lgQGK1pWfVhLGZEuPoa4LzBp8A6f2kf4wO2UTcwmCosObmBuxTqZbW7FNRCu6VFdfS1Xwrai9hVlsUs3GC5ydKOZnP4MX9kk5pHQXT9VXZJhMeDntnceq+XGgtaW45QGpbs9ibO+RLrHwzq9CY6xp158WnxOsfCCVqdTM130xt29lxhgVBabrqFrt5W735mqfBetTapRqzaG28+eaHdyHyO1C5TpiSswdl96AUjxbvKWeZoOhdZJd7Ck0EbVr3nbAVRF4YVZx62K0EfuBU6mHLr7TC0He7oet02t4DMIpKUjeczd2DP/5sOn/LX/RvvOlPF2X3hZinNi9qYrsQLt70u97L3vQf5FgLUbDS/arXa/iNdrXS8Fu9StBp1yuNqNqudKpRrdPrRGG90XtgoaMCHLT8KKh265WqG0WVoOrkS6k3KrXA81pBrVXvBq0Hc7uDFRbXhakLjjt/AVBLAwQUAAAACAAROrZcoCM6fE0KAAAZRgAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbK2c224bORKGX0XQAHMXi4cqHjK2gSSLxe7FLoIEM3vdsduxMDqt1I6TefqtbskW1WAXi8gGgS3J6qr6SXbx46mvn7f7Pw+PbdvNvq9Xm8PN/LHrdm8Xi8PdY7tuDlfbXbuhvzxs9+umo7f7r4vDbt8298NF69XCKOUW62a5md9eD5993N9eb5+61XLTftzPDk/rdbP/8b5dbZ9v5nr+8sGn5dfHrv9gcXu9a762n9vu993HPb1bvFq5X67bzWG53cz27cPN/J1++yGY/oLhG38s2+dD8nrWNV8+t6v2rmvvB09/bbfrz3fNqv13H/yKPlNqPusFf9lu/+wv+Sd9UfVxD5f1jhr69a390K7o6+8tRfvfwfV7+/Y9uvlrdP3F6euXOP4+FBPJ/tIc2g/b1X+W993jzTzMZ/ftQ/O06s6fxdfPPm2f/9GeisNcDV7utqvD8HP2fPy2hfns7unQbdeny/uS7H6s2pu5oZfr5Wb4aN18P5Vpeq260jEaHTxOGvEnG+Zkw4xtCPzb07V2dK2OV07qH042YGQjFq/E05VYfaV7KTdnw+B3cSz+oUb/1nTN7fV++zzbD9f2tWTxyrxKea06skuF0Tenu/67704fUInSX5eb/n743O3pz0uy3d3++gv4aNxvv/6CWhlFvyEYB/17CKB+u150FEn/3cXdyeT7o8kwaREtuDhY0r63DNEayFn6cArOTptSwYXBBA4mUWsVC0EuqJhey8ocy0r7KygWlTl+ABPRgNLG6mCMtxZ0NLmiOZrwg4U+FX27hevFt1TyVEW4GPQbRAxvorOGUWQrFFmBIq8CUHFatCqnyDJapuot1QKW0QIVWkCgRZEcDQboDsppAUbLlN1UCwKjBSu0oERLDCboCOAxpwUZLSjQ4pDR4iq0OIEWG6yHaNA7ndPiGC1OoMU7Rouv0OIFWhA0KEUtTdmcFu7e9wItQTFaQoWWINGiUfUZ1rqQ0xIYLVPpP9USNaMlVmiJksyM6FTQwcbs/RIZLbGsxSsuJ2tVIab/8s+q6W1Myjn98Wf06Bo9+v+gR3N69E/rqcEALeKAgh7D6REwQEFPDQRoCQWU9HAYoAUcUNBTAwJaQgIlPRwKaAELFPTUwICW0IBHBY7gJvhsptYcDmgBD3jNcZquAQItIgKg3icQ3tt8/XBIoAVM4Md2L/XUQIGWUEH/3xqnosrr4bBAC7jAR449dQ0YaAkZWKvQREQa62T1cGigBWwQxuV0qacGDrSEDpQNEQI6M3H/cHigBXwQNMegpoYPjIQPPI3W0VMdYZaoDccHRsAHwXLsZmr4wEj4oKSH4wMj4IOCnqppAgkfUCJwNH5zKt/eDMcHk5M1qR7k8rWp4QMj4QPlaCSPGEPIz3twfDA5v5PqGY9yL/XU8IERzRQU9HB8MGlZrqeGD4yEDwAjjUkVwHiO4qSH44OSg1RW5IanpgYTjAQTvNXBAf3MT04ZDhNKDgZZDqN/ozGw2buGFoyEFgy1PBp6Rwv57MBOIBYcJLUVFZv0aqDBSKDBAYAPwUedl8VBQ8lBWlt8I6xhByOaWYikygSIKi+LY4eSg7S2NDvtW4MQVoIQ0VpNHa5TOouslkOIkoOkthzbCG0NSVgJSUBQ1hsa+bksuVqOJEoO0toCLsHbGqCwEqCw1Dy0C5ZIKSuLA4qSg1QWcpnQVi0+iLiCBtLoKMermJXFLj8UHKSyxqV2KasGL6wEL8Chs0E57/OyOLwoOUhljenlUlYNZVjZLESkZKi88VlqshxllByksiIHt7aGMqyEMqjbotGh0ta6rCyOMkoOXmTFN0opbgxvayjDSijDgdZaBa8hv7LHUUbJQSrLsCmjhjKsaNFCRYR+3RLymZCjjJKDVNbY/KWsGsqwEsowXlONRadNvjvmKKPkIJWFXMqAGsoACWUYpL4YnFfjRbrTkixHGSUHqaxxb38pq4YyQEIZaIjekZp+fvwIHGWUHKSyPJcyoIYyQEIZaCBgMNGNB0QnWRxllBykTMiiLtRQBkgog3qV4DT0/7KyOMooOUhrK3KZEKp2O0gog8Yl/RDZxJAdmAC730FIGSRLK268BTWUAaKdDwG1cZZSVb4RcpRRcpDKMmwjrKEMEM1lGNToqDHmN3QARxklB+m9ZdmUUUMZIKGMkiyOMkoO5LJqKAMklFGSxVFGyYFcVg1lgIQySrI4yig5EMvCGspA0XIILws5yig5kMuqoQwUrYoUZHGUUXIgl1VDGSihDO9BoY9I49esLI4ySg7SBD9eu7yUVUMZKKEM1098Go0+P5eBHGWUHKSyxhBzKauGMlBEGSRJG4KAPOoiRxklB2kjZGeesGp7pYQySrLYDZYSyhDJqqEMlFBGSRZHGSUHclk1lIESylBBeVTOW5edy0COMkoOUlnsPCHWUAZKKEPHSAN/r11+aQE5yig5SFOGY2urhjJQQhkAlN2902h9VhZHGSUHqSzPDUxcDWU4CWXYEBTpis5kG6HjKKPkIJUVuEboaijDSSjDAAaL1Pbzo2PHUUbJQSprfGTiUlYNZTjRFk1jEBWJG2+dOsniKKPkIE0Z7GS1q6EMJ6EMiwG0NVP3luMoo+QgqS0z3t5+KauGMpyEMpx2JoKOarxp/CSLo4ySg1SW5uDJ1VCGk1BGINDtJz+jyTdCjjJKDlJZ40njS1lV5zkklIFggonRuYlGyJ7okK+YmPGCzKWsGspworMdUVt0w9p4VhZHGSUHqazxgsylrBrKcKJ9GdT/u/78zfgw1kkWRxklB6ms8d69S1nxHPTUhsrjLaMjAYSJaqI74uBh0nIapuP6IZ+cBpraJ3kcxAYFNNpTMb8M5TkYmLSchum5BOz1Ocyp7Y+vx/sC6uAwe6d6rnOftJyGOT5qdxmmOYc5tavxtL0BgwOIEbMzHZ7rrCctJ2FaxZ4Rs+cwpzYrHkuTUInuJ2vy2+w81/lOWk5gwrNLxB7OYRYmWAKgD/0u1GyYXGc6aTkNc7xl6DJMPIc5dfTgCJ4OQ0QVML+dwHOd46TltNI1l2e9O4c5daLgJUzUOniT31Hkuc5u0nIaJnve1/tzmFMHBY4jfzIZADTdRdkw2ZOJpSMIfZjsUV4fzmFO7f8/dkY2IA1Ijcp3Rp7rjCYtp2Gyp3T9uRfybC8ENNpD8NqOx5inMLleaNJyGiZ7aDUk5zjZXija4Pry9CGb3gPXC01aTu/08cbbyzDPvVBge6FADELYY1x+IjtwvdCk5bQ0+WOz514osL0Q3T2a8AAgf348cL3QpOU0zPyJ2EXyjIe77eZ+2T8EpFkdH+PRLTdfXx4A8k6/facJO/CYux4+Pa3aWfdj197M7592q+Vd07V/NKun9jCf7fbL7X7Z/Rig8f77w/EpI8fnSmRcHJ9/8q9m/3W5OcxW7QNpUVeUFvZH8hxed9vd8Ipy75dtR1z68u6xbe7bff+OOrmH7bZ7eXN+rsrTbkbxtJuu6V3fzHfbfbdvlh1F2uza/eflX21/KLgP8PVpMLf/A1BLAwQUAAAACAAROrZcqBORybcOAADcXwAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQyLnhtbKWcbZPTRhLHv4rLqaJy92J3nh8CSxVsFkiAsIc3d7w1IMCV3bVjm0Dy6W+k0YjRzrRanbzBa2OP/hpJ/Zvu6e4HX7b73w6fmua4+HpzfXs4W346Hnc/nJ4e3n1qbtaHk+2uuQ3/82G7v1kfw9v9x9PDbt+s33c/urk+FYyZ05v15nb58EH32eX+4YPt5+P15ra53C8On29u1vs/HzfX2y9nS75MH7zefPx0bD84ffhgt/7YrJrjr7vLfXh3OozyfnPT3B4229vFvvlwtnzMf7jgirW/6L7y303z5ZD9vWjP5e12+1v75qf3Z0u2bMe+bRZ/rnbXm+5oi+N296L5cDxvrq/Plo/EcrF+d9z80VyGr50t326Px+1N+/9B53F9DB992G//am67YzbXTfhuULMrvhwHiYM+bk/y917xcjihVlT+d1L+pJvZMFNv14fmfHv9v83746ezpVsu3jcf1p+vj98+88Nnr7dfnjX9DOoT0x7l3fb60P27+BK/rU6M94I7q5eLd58PQWs/TKvv+Od10/11s7mNr+uv/dXIhmjHnjuG6McQd8fg4E9F/1PZ/1Te+alHD6r6XyryL3U6ZSNd9+vTOIHdNflxfVw/fLDfflnsu1+08yz1iRhmYZj8MG44j/ZGe9d+93H8IJxQ+M/NbfsMrI778L+bMPTx4b3vtFTG37/3nXLcmvbVS6HaV2uVv//g9BiEtN89fdePeP53R9SSxfdGMlYb+SKOrICR/73IxmJO2/ZVGR6P4VynWtvwfjEhgjMpJr8wUtnOgwojtz9QrvtBGMGo2qGV9t3IYUTluhG881GTcfHVm/sn4xM/Ddd0uLAiXlhuTgR6XUX3gQbm6vzn9nhO9EIF687MOB3nLuipXdk4punGbK3nHw/Vg9M/KkIlQaicFJrJ40aL+3cvTE2mLGTqukxFkKkwmWHWutkUisXLLb2dJVcVck1dribI1YhcZbk33awKOW9WdSHT1mUagkyDyhRcd8+Gl2aWTFPIdHWZliDT4g9TutxcCNcJ9aoT7rWpm0lbCPV1oY4g1E0KvXi5qihNN64SdbN77gqlnNeleoJUPyn16dWbOVfbl9JkXRpnBG3tlyfEPX/6eI64fpSROsBgck5Rx+eo6+8/pbWc9eD0o47UAnaTU0DEp0n0/Co3RiaS0TtRJxAvEcQBa8kpEOLTFHrxtHteWL9CcJ1I6zvLGRYAChBbgogDNpNTUMSnWfTqfPX81euLR1VJJWw4YB85BTd8mjerc9JFLlHDAdvIKbDh07T5svs66zkpCSMYoI7CGD4NmfZRjutGzUy3bpwhtWSMACw3p1CGT2OmlSoE75a4TPF5UkvICAFIpVCGT2Omk6q7ZboNnvo8qSV0BAAdQYGOmIZOK1XFxZC1yZmZeSOIkkQCIJGgkEhMkyg5PZ1EVV9ciBI7AsCOIPk/09hppTmlstmcMYslfAQAH0GBj8BcIKUFi8hhbN49Kkr0CAA9goIegblBKmhz/YLdTlp5UaJIACgSFBQJ1PXRxuV+b91TmwN4UYJKAKASFFAJ1C3qXfc2tuDm3RAltSRALUGhlkCplaIMyc0gL6NESTEJUExQKCZQivX0Ui6InjfLJcUkQDFBoZhAKZakehX9ZM5iTEkxw6JkAdjfkmYSoJmk0EyiNEtrA6ZtZymsjvdyWB/WPWZZUkxCASgKxSRKsbR0ZaafZSfmPXaypJsE6CYpdJMo3e4ETlqJMSDqTJxlwecFVGRJPQlQT5Lifij17p6Cd5F+igNrCVlSTwLUkxTqSZR6Vkcj5mKMMt0w34LHUBRIlhSUAAUlhYISDwCmwIDUJo9r96cSHksgHCRL6kmAepJCPYkHA3WHkGA5IrhRdMiSegqgnqRQT6LUczw+VtaL6Oyannbc8LrUknIKoJykUE6ilHPCJ2SMjJvq72nLO/rBKzlZUk8B1JMU6kmUemkvSrH+nrXh8atKLCmnAMopCuUUSjkvO7cygFnb+8+uns3aEShBpwDQKQroFAq6kdq/s1hTJfAUtPtCAZ5CgddLnx8dUZXtLABsigI2hYKtn91vRgGd1RJsCgCbIm1qoWDrjUFrelm2rpwxuyXYFAA2RQGbwsAWsBUZPPJAv19d/Hjx5l9VpSXPFMAzReGZwngWEZutHoz3sj6ZJcc0wDFF4ZjCOJYkaq5mbr6qkmMa4JiicExhHEt+xPCKSy25pQFuKQq3FMatYKNEZ6P6veLo8kxKLfmlAX5pCr80xq8g1YoYLJHz9pF0CS8NwEtT4KWxXa+O/3HvffDX0p6887YutsSVBnClKbjSGK6C1D6NhvN8V2TG/JbY0lDCAAVbGk3E6HNqEgvmb+joSkYGgC9NwZeem5QxODX47JbY0gC2NCkfA8fWXalYAoEuuaUBbmkKtzTOrXqqSzARcdenC/pXJZccMwDHNIVjGucYkJ2T4mSOA56DLnlmAJ5pCs80zjNno98lw2O1+CdrMV2CzgCg0xTQaRR0vA8/ahVv6pGpmCW9BJ8BwGco4DMo+Hrp8xOkTAk+A4DPUMBnMK9tkCplHx/xkdldJmJVaok9A2DPULBnUOwNUn2/+caYmmSHKXFnANwZCu4Mirv0eM1eT5oScwbKkaNgzqCY693zwbqhYUdTSToEMGcomDMo5njMDGgdCzXL9zUl5gyAOUNKPUQx1yfIpfzcNvoPSCyxZgGsGQrWDIq1wS2LwdsZ92iJMwvgzFBwZlCccWVlvPBiXkqIKallAWoZCrUMSi0hHQ2wpqSUBShlKZSyKKVS1H4U7Jgj2Za0sgCtLIVWFqWVlCzfqlTOsmn3wZa0sgCtLIVWFqVVmMUYChd8FFpgcRMtBvCqkktqWYBalkIti1IrgbVfAygtzTx62ZJeFqCXpdDLovRSTKQ1uBnWt1WJJbUslN1NoZZFqdXlRt/rq0Mou6e2kjYP0MtS6GVReqXykyF3qPfWccmVFHqAZpaURI/SrEjIQXPoS5o5gGaWQjOL0ixJTdvq+KyWNHMAzSyFZhal2SDVSj5vVkuaOYBmjkIzh9LsbqlPv1gcsgj75IW4aKyWKpRUcwDVHIVqDqVam2FzP9s5Q2fZlVRzANUchWoOpZpycZNPCzEv8dGVNHMAzRyFZg6v/coveObYaMfqnq0rKeYAijkKxRxOsRggn5/y7EqaOYBmjkIzh6Z+9NvlrZ1NAdxxqSN4z5Y0c1DVEoVmbppmbWx/frTZlfzyAL8chV8ODzJ6NXYY0/4e5uO6SjEYwDFHKgfDUxxTFauO2w8pPjPKdKtKLnnmAZ45Cs/cNM9+XV2uqnJKZnmAWZ7CLD87Tihc3Hbu6gmqtWolmzzAJk9hk59m0+rFvEq6EkcesEaegiOP74Tl9YdzrJAvceQBK+QpOPLTOHry4klKPsyyopww81LyfaUajAF2yVPQ5KfR9OOzF1U5lUowBpV1UvDjp/HzpM3DqAqqVH0xwKJ4Clz8NFyerl7+cvXql6qkkiTBAQQkUVDip1Fy8XL1fbeiiK68l91ScrykrKay+JIjnEE2hgISP8Mh4iYHCP48VMqKGbAA9qTCYqSyeHW5PRzfbmcV+/lKgTGDCk4ZrcR4miznv/6TsEMafSwcKj5lpOpjFj/yUxmlMRqd3LmUnimNmd7sSWO3ldOZbqhClZHqkBmeQpjq6b5tTtWTddNY4+mFalQZqRKZTVPo0flPi4s3l68vVtV1UPr5uFofqk9lpLJjhmcL9s4vczH9XWmV5+O0sHSz1krpWOPTAEv7dSYS9X7ubqMPV3t2fTIrUeWhMEMwF6Q5RpPghxIDrXkulyC/gjUo8zUYDZJ8PMbnYrlBcOFnFjGzinMEJb8FW0GSi/tHPNVPRdk6Lk6Df8RHhR7a8TyfZM5pVRwoKPcs2BbSaeE7XPEBHF67bLSr8jzwxXg62Pg8oLuJ1oUDacPRedy6z07sWyrJvkQoLp3SRk0bnLHR4Nh6IiCvNOvw0O495zwTiSbVt5UUMXrJZhbDV3pxeGh/PvghmRhyXRi5bw2v9ODw0I58sNmZuBkpF6y/XH3rLJuSTxkkpsSdh/beOVeZGLymq2ZQu8vmaDmFvNJ1w0N7a5xnMEPaasAi057aALegtb6nxivdNjy038O5ycShW1DC+5SHxPJ7LL1XDCgx45UmGx7aLuHcZqKmAfT0akVCZaWDhof2Qjh3mQ6ULMN+HLoorvTG8NBWAec+E4HvDdlYgZLqyYsFG7Y3zyvdMDy0OcBHbYHQ0qm+7FpZBWRvpzHGB4ceLJEZbbxxxd3GY8LJb/s7dTG1CBpktEVmtGe0qoiNP7R0fdIUcxoQEcdqAZKpgKyzyKwz3oUipRQMOSZp4Zzno9ZVyYoqKMLNRWamkYYTj84vFhdfd/vmcKgfWNUODBkSkZlevI1EQdKas1NXpWuqILMiMpuLdoaIqOxCB7ZeS5TGuHN0yH0RmXHtOzG0CwDEuYr9XkimNg0+lgVakszW9l0XpmSN2tDMYnYadCwHtC2Z1UWaNTx/+uLe0CYlPUOtxyyR+fE1QZB9kXlvM4bNTyAx772yuzW/rr7POgw6lgMZGpnZ3r4nAijnPxPPc/ptftSwaISWKzKzskjLhBdvFpePfrl6VY+qyIpZDceFLJjM7CrS5yAc9+fLy1erq/pxK4YzHBcyYDKznHjTglTs2/WdvfetHFUN0caYfYH7gLJiZoNMyKLJzM6ijQqCbyrHz+7cRgVp7LuywIaBmaFFmxEMpQ2jMJJQ/yhqKyumOQiGrI7MbHPfPwCOyhbZVUz3jR+ivLDM6nOI2xOpy+tNtBzLA21QZqORNgTD0xceBlk/tqsdG27+mBnkvvifQ42j84P7+sH7EfT44NCJq7xXI5s+eC1twApiRC8d5I6+u9b4NOsV3rawf7nef9zcHhbXzYfwM3YSLu8+xmK6v4/bXfdXGDW2jk/vPjXr982+fReuxoft9pjefGuN/3m32O43ze1x3TagP1vutvvjfr05Lhe79a7ZrzZ/Ne3N2rYxHxr6P/w/UEsDBBQAAAAIABE6tlwDiiMZ0AQAAAUnAAANAAAAeGwvc3R5bGVzLnhtbO1aTY/bRBj+K5ZP9ED9lTg2JJHAIhISSBXtoRJFyInHyUj+CPZklfS07QkJoapSC5WgEgekXrh1Dxz4RST7H5jxOI6TzJtkN96sF3C065l5v573Y2bicdopmQXo4QghIk3DIEo78oiQ8UeKkg5GKHTT+/EYRZTix0noEtpNhko6TpDrpUwoDBRdVU0ldHEkd9vRJOyFJJUG8SQiHVkrhiR++9yjg2ZDlrg6J/ZQR1a/vffx118h75snH6hP7slKt63kerptP45K6gyZj1CtboikMzfoyPN3Lxdvz6W/3/82f/WaCQ/iIE4kQtEjhoCOpE85q8Z7zLFcOMRRnGQWudr9ykduktJYcX26zcZ8N8TBjA8ZKwTJsN+Rez01u8owrFOiEMZBPQ2Cfk6uHM2LHxa/nFeVDl4VkKnnb+bfv6o48lXa2o5x7ih39XDjx1e6kV23UWjrARdkd8I7KY6GAboNTIJ16ChMjhvgfoI3LetXnXZbIK5h77rlfvJVLbuxTQUHQbGpNGU+0G2PXUJQEvVoJ5PJBrdIUt5+NBtTBMPEnWl6Uz5YII0D7DGTQ6fsiS5LBDM4H6r3Vdu2G5amWU3DsGzVyrzs5/xby5hSMnUkCPVUhja9bVmm0VIbhmGa1/E2u9G89uPEQ0mR2Za8HOq2A+QTKp7g4YjdSTxmVmJC4pA2POwO48jN0r6UKEtK2bekjkxG2bccXoPuhMR5CSqMKde+lzfjyiDsZaU8S5R7eTlbtb4otwdTucvh3C7bfQWyLbHHr22Bfd5tS9xpH687bXfi3lHwNx3wvEGXsQEKgodM32O/WMsaVOvULz1FqewZKiqadAHMm1xN3nHH42D2SYCHUYiWD2TusiudoYTgAdspB7SL+FY59WFLjZNZsk5mya7cklJOIE9nKZN6s6pUSmN8FpNPJ7Smoozlu0lM0IME+Xia9ac+lz0qPvp/EdIoTvBTaq0E6uowjRXMRhlmo14wS9E0bj6abLE9DqR6M7GsIaRK0tus3/w1V5D0MiStXpGDEtysX81VB6mSyLXuxtTQariqQKGrWYahyJn1i1wdVpX8ncfeL3mHolJ26YeWV6NugajL1nRQ9PQaRO+wJ4VKiuj/3BxT2XdmClZTRkr+jFs6u1g7uShGJXb835EvX/66+P18pULqT3BAcFQoFAtI+lJEEzK9frN4e3H541/zd88un/2x5NXX1FuyNMKehyJAyeLni8VPF9L8/Z8SoM9Y02dv6Fs9/NNoeFN/7V24x5LB30WsHRHZjqqqZvldyeax+taReK/ntJzP4CPxzBT7T2EQtx+g9azQmHvIdycBeVQQO/Kq/SXy8CTUC64HrBhyrlX7C3ZmppnFexlqC0cemiLPybsUbQm2qq4OwzYpvewSUyAZThNTGA2yAyGAZLgUZOff5I8F+sNpEDZLSLFAGQuU4VIiipN9IDtiGZteYk9tm71ugiLqOEIEDhQ302R/Ym0QNiYB2WGWrhZrONtwheyuAyinuyoE8hSuRMhTONaMIo4bk7BtcbYhO0wCygJUO8y+2A6rKbGMYbCsQtigGQxTbBuisFoU16hpAtEx2UecH2iWGIZtiymMJkbAfh8iprDZCFMgBAwDROG/RFE29iNluU8pqx++df8BUEsDBBQAAAAIABE6tlyXirscwAAAABMCAAALAAAAX3JlbHMvLnJlbHOdkrluwzAMQH/F0J4wB9AhiDNl8RYE+QFWog/YEgWKRZ2/r9qlcZALGXk9PBLcHmlA7TiktoupGP0QUmla1bgBSLYlj2nOkUKu1CweNYfSQETbY0OwWiw+QC4ZZre9ZBanc6RXiFzXnaU92y9PQW+ArzpMcUJpSEszDvDN0n8y9/MMNUXlSiOVWxp40+X+duBJ0aEiWBaaRcnToh2lfx3H9pDT6a9jIrR6W+j5cWhUCo7cYyWMcWK0/jWCyQ/sfgBQSwMEFAAAAAgAETq2XLKZAN+bAQAA/gIAAA8AAAB4bC93b3JrYm9vay54bWy1kk1u1EAQha9i9Z74B2mUGY1nQwSJFEFEULLuscvjUvrH6q4ZJ1kh4AIRLBNOwGKQ2HIh4tyBcltORkJCbFh1vVdW1/NXPW+tu1haexFdamV8LmqiZhbHvqhBS79nGzDcqazTkli6VewbB7L0NQBpFWdJMom1RCMW8/GuExcv5n1xhtD6J7+X0QY9LlEhXeUi1ApEpNGgxmsoc5GIyNe2PbQOr60hqU4LZ5XKRTo0zsARFn/Yp32ed3Lpg3N5jqa0LdfTZMo3Xo36WZqybIM6x5LqXDzfn2SP3iHgqqZcZPwdmySXbyWhzcWklxU6T2FSyCkLwg3w0EGtyb5EReAOJMErZ9cNmlUfh2nEOzgCuvEcuM/cv5C3VYUFHNhircHQgN6B6gMaX2PjRWSkhlx0X2+6T99/bd93dz/uP99GmzTZC1x54FE5MCMOubMBN0NuuKMyxP1/0R4+3t1vt92Hb1H38wtne7i53QmW/SVYFjiO8Eqo0ED5mi/17PODKE5c1B/hB6dpkk15YWulXrD3xhxbWY67GN/j4jdQSwMEFAAAAAgAETq2XI33LFq0AAAAiQIAABoAAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc8WSTQqDMBBGrxJygI7a0kVRV924LV4g6PiD0YTMlOrta3WhgS66ka7CNyHvezCJH6gVt2agprUkxl4PlMiG2d4AqGiwV3QyFof5pjKuVzxHV4NVRadqhCgIruD2DJnGe6bIJ4u/EE1VtQXeTfHsceAvYHgZ11GDyFLkytXIiYRRb2OC5QhPM1mKrEyky8pQwr+FIk8oOlCIeNJIm82avfrzgfU8v8WtfYnr0N/J5eMA3s9L31BLAwQUAAAACAAROrZcbqckvB4BAABXBAAAEwAAAFtDb250ZW50X1R5cGVzXS54bWzFlM9OwzAMxl+lynVqMnbggNZdgCvswAuE1l2j5p9ib3Rvj9tuk0CjYioSl0aN7e/n+IuyfjtGwKxz1mMhGqL4oBSWDTiNMkTwHKlDcpr4N+1U1GWrd6BWy+W9KoMn8JRTryE26yeo9d5S9tzxNprgC5HAosgex8SeVQgdozWlJo6rg6++UfITQXLlkIONibjgBKGuEvrIz4BT3esBUjIVZFud6EU7zlKdVUhHCyinJa70GOralFCFcu+4RGJMoCtsAMhZOYoupsnEE4bxezebP8hMATlzm0JEdizB7bizJX11HlkIEpnpI16ILD37fNC7XUH1SzaP9yOkdvAD1bDMn/FXjy/6N/ax+sc+3kNo//qq96t02vgzXw3vyeYTUEsBAhQDFAAAAAgAETq2XEbHTUiVAAAAzQAAABAAAAAAAAAAAAAAAIABAAAAAGRvY1Byb3BzL2FwcC54bWxQSwECFAMUAAAACAAROrZcogrxBzEBAAB7AgAAEQAAAAAAAAAAAAAAgAHDAAAAZG9jUHJvcHMvY29yZS54bWxQSwECFAMUAAAACAAROrZcda6PvKQGAAADIQAAEwAAAAAAAAAAAAAAgAEjAgAAeGwvdGhlbWUvdGhlbWUxLnhtbFBLAQIUAxQAAAAIABE6tlygIzp8TQoAABlGAAAYAAAAAAAAAAAAAACAgfgIAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWxQSwECFAMUAAAACAAROrZcqBORybcOAADcXwAAGAAAAAAAAAAAAAAAgIF7EwAAeGwvd29ya3NoZWV0cy9zaGVldDIueG1sUEsBAhQDFAAAAAgAETq2XAOKIxnQBAAABScAAA0AAAAAAAAAAAAAAIABaCIAAHhsL3N0eWxlcy54bWxQSwECFAMUAAAACAAROrZcl4q7HMAAAAATAgAACwAAAAAAAAAAAAAAgAFjJwAAX3JlbHMvLnJlbHNQSwECFAMUAAAACAAROrZcspkA35sBAAD+AgAADwAAAAAAAAAAAAAAgAFMKAAAeGwvd29ya2Jvb2sueG1sUEsBAhQDFAAAAAgAETq2XI33LFq0AAAAiQIAABoAAAAAAAAAAAAAAIABFCoAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzUEsBAhQDFAAAAAgAETq2XG6nJLweAQAAVwQAABMAAAAAAAAAAAAAAIABACsAAFtDb250ZW50X1R5cGVzXS54bWxQSwUGAAAAAAoACgCEAgAATywAAAAA"

def get_template_wb():
    data = base64.b64decode(TMPL_B64)
    return openpyxl.load_workbook(io.BytesIO(data))

def save_playauto(rows, out):
    wb = get_template_wb()
    ws = wb['작성가이드 v10.0']
    ws.delete_rows(2, ws.max_row)
    for r in rows:
        ws.append([r['묶음번호'], r['택배사'], r['운송장번호']])
    wb.save(out)

def build_order_map(sku_df):
    m = {}
    for _, r in sku_df.iterrows():
        for oid in str(r['쇼핑몰주문번호']).split():
            m[oid.strip()] = str(r['묶음번호'])
    return m

def extract_cj(wb_df, order_map, seen=None):
    if seen is None: seen = set()
    rows, new_seen = [], set()
    for _, r in wb_df.iterrows():
        if str(r.get('상품명','')).strip() == '분리': continue
        bundle = None
        for oid in str(r['고객주문번호']).split():
            if oid.strip() in order_map:
                bundle = order_map[oid.strip()]; break
        if bundle and bundle not in seen and bundle not in new_seen:
            new_seen.add(bundle)
            rows.append({'묶음번호': bundle, '택배사': CODE_CJ, '운송장번호': str(r['운송장번호'])})
    return rows, new_seen

def extract_kd(kd_df, sku_df):
    name_map = {}
    for _, r in sku_df.iterrows():
        if str(r['SKU상품명']).startswith('(경동)'):
            name_map[str(r['수령자명']).strip()] = str(r['묶음번호'])
    rows, seen, unmatched = [], set(), []
    for _, r in kd_df.iterrows():
        name = str(r['받는분']).strip()
        bundle = name_map.get(name)
        if bundle and bundle not in seen:
            seen.add(bundle)
            rows.append({'묶음번호': bundle, '택배사': CODE_KD, '운송장번호': str(r['운송장번호'])})
        elif not bundle:
            unmatched.append(f"{name} / {r['품목명']}")
    return rows, seen, unmatched

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("페이퍼팝 송장 전송 자동화")
        self.geometry("680x720")
        self.resizable(False, False)
        self.configure(bg="#F5F5F5")
        self.sku_var    = tk.StringVar()
        self.wb_var     = tk.StringVar()
        self.kd_var     = tk.StringVar()
        self.prev_vars  = []
        self.output_var = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        self._build_ui()

    def _build_ui(self):
        BG="#F5F5F5"; CARD="#FFFFFF"; BLUE="#1B4F8A"
        LBLUE="#E8F0FB"; GRAY="#6B7280"; GREEN="#16A34A"; BTN_FG="#FFFFFF"; PAD=14

        def section(parent, title):
            f = tk.Frame(parent, bg=CARD, bd=0, highlightbackground="#E5E7EB", highlightthickness=1)
            f.pack(fill="x", padx=PAD, pady=(0,10))
            hdr = tk.Frame(f, bg=BLUE); hdr.pack(fill="x")
            tk.Label(hdr, text=title, bg=BLUE, fg=BTN_FG, font=("맑은 고딕",10,"bold"),
                     anchor="w", padx=10, pady=6).pack(side="left")
            body = tk.Frame(f, bg=CARD, padx=10, pady=8); body.pack(fill="x")
            return body

        def file_row(parent, label, var, cmd):
            row = tk.Frame(parent, bg=CARD); row.pack(fill="x", pady=3)
            tk.Label(row, text=label, bg=CARD, fg=GRAY, font=("맑은 고딕",9),
                     width=14, anchor="w").pack(side="left")
            tk.Entry(row, textvariable=var, font=("맑은 고딕",9), state="readonly",
                     readonlybackground=LBLUE, relief="flat", bd=1).pack(side="left", fill="x", expand=True, padx=(0,6))
            tk.Button(row, text="찾기", command=cmd, bg=BLUE, fg=BTN_FG,
                      font=("맑은 고딕",9), relief="flat", padx=8, cursor="hand2").pack(side="right")

        hdr_frame = tk.Frame(self, bg=BLUE, height=54)
        hdr_frame.pack(fill="x"); hdr_frame.pack_propagate(False)
        tk.Label(hdr_frame, text="🐸  페이퍼팝 송장 전송 자동화", bg=BLUE, fg=BTN_FG,
                 font=("맑은 고딕",13,"bold")).pack(side="left", padx=16)
        tk.Label(hdr_frame, text=datetime.today().strftime("%Y-%m-%d"),
                 bg=BLUE, fg="#A5C8F0", font=("맑은 고딕",10)).pack(side="right", padx=16)
        tk.Frame(self, bg=BG, height=10).pack()

        s1 = section(self, "  📂  오늘 파일")
        file_row(s1, "SKU 매칭명",     self.sku_var, self._pick_sku)
        file_row(s1, "운송장출력데이터", self.wb_var,  self._pick_wb)
        file_row(s1, "경동 발송자료",   self.kd_var,  self._pick_kd)

        s2 = section(self, "  📂  미발송 운송장 파일 (선택)")
        lbf = tk.Frame(s2, bg=CARD); lbf.pack(fill="x")
        self.prev_listbox = tk.Listbox(lbf, font=("맑은 고딕",9), height=3,
                                       bg=LBLUE, relief="flat", bd=0,
                                       selectbackground=BLUE, selectforeground="white")
        self.prev_listbox.pack(side="left", fill="x", expand=True)
        sb = ttk.Scrollbar(lbf, orient="vertical", command=self.prev_listbox.yview)
        sb.pack(side="right", fill="y"); self.prev_listbox.config(yscrollcommand=sb.set)
        br = tk.Frame(s2, bg=CARD); br.pack(fill="x", pady=(4,0))
        tk.Button(br, text="+ 추가", command=self._add_prev, bg=BLUE, fg=BTN_FG,
                  font=("맑은 고딕",9), relief="flat", padx=10, cursor="hand2").pack(side="left", padx=(0,6))
        tk.Button(br, text="삭제", command=self._del_prev, bg="#9CA3AF", fg=BTN_FG,
                  font=("맑은 고딕",9), relief="flat", padx=10, cursor="hand2").pack(side="left")

        s3 = section(self, "  💾  저장 경로")
        file_row(s3, "출력 폴더", self.output_var, self._pick_output)

        run_frame = tk.Frame(self, bg=BG); run_frame.pack(fill="x", padx=PAD, pady=4)
        self.run_btn = tk.Button(run_frame, text="▶  송장 파일 만들기",
                                 command=self._run, bg=GREEN, fg=BTN_FG,
                                 font=("맑은 고딕",12,"bold"), relief="flat", pady=10, cursor="hand2")
        self.run_btn.pack(fill="x")

        s4 = section(self, "  📋  실행 로그")
        self.log_text = tk.Text(s4, height=10, font=("맑은 고딕",9), bg="#F9FAFB",
                                relief="flat", bd=0, state="disabled", wrap="word")
        self.log_text.pack(fill="both")
        self.log_text.tag_config("ok",   foreground=GREEN)
        self.log_text.tag_config("warn", foreground="#D97706")
        self.log_text.tag_config("err",  foreground="#DC2626")
        self.log_text.tag_config("info", foreground=BLUE)

    def _pick(self, var, title):
        p = filedialog.askopenfilename(title=title, filetypes=[("Excel 파일","*.xlsx")])
        if p: var.set(p)

    def _pick_sku(self):    self._pick(self.sku_var,    "SKU 매칭명 파일 선택")
    def _pick_wb(self):     self._pick(self.wb_var,     "운송장출력데이터 파일 선택")
    def _pick_kd(self):     self._pick(self.kd_var,     "경동 발송자료 파일 선택")
    def _pick_output(self):
        p = filedialog.askdirectory(title="저장 폴더 선택")
        if p: self.output_var.set(p)

    def _add_prev(self):
        paths = filedialog.askopenfilenames(title="미발송 운송장 파일 선택",
                                           filetypes=[("Excel 파일","*.xlsx")])
        for p in paths:
            if p not in self.prev_vars:
                self.prev_vars.append(p)
                self.prev_listbox.insert("end", os.path.basename(p))

    def _del_prev(self):
        sel = self.prev_listbox.curselection()
        if sel:
            idx = sel[0]; self.prev_vars.pop(idx); self.prev_listbox.delete(idx)

    def _log(self, msg, tag=""):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg+"\n", tag)
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.update_idletasks()

    def _clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0","end")
        self.log_text.config(state="disabled")

    def _run(self):
        if not self.sku_var.get() or not self.wb_var.get() or not self.kd_var.get():
            messagebox.showwarning("파일 미선택", "오늘 파일 3개를 모두 선택해 주세요.")
            return
        self.run_btn.config(state="disabled", text="처리 중...")
        threading.Thread(target=self._process, daemon=True).start()

    def _process(self):
        try:
            self._clear_log()
            today = datetime.today().strftime('%Y%m%d')
            self._log(f"[{today}] 처리 시작", "info")

            sku_df = pd.read_excel(self.sku_var.get(), sheet_name=0)
            wb_df  = pd.read_excel(self.wb_var.get())
            kd_df  = pd.read_excel(self.kd_var.get())
            self._log("✔ 파일 읽기 완료", "ok")

            order_map = build_order_map(sku_df)
            cj_today, seen_today = extract_cj(wb_df, order_map)
            all_seen = set(seen_today)

            prev_rows = []
            for path in self.prev_vars:
                df = pd.read_excel(path)
                rows, seen = extract_cj(df, order_map, seen=all_seen)
                all_seen |= seen; prev_rows.extend(rows)
                self._log(f"  + 미발송 {os.path.basename(path)}: {len(rows)}건")

            kd_rows, kd_seen, unmatched = extract_kd(kd_df, sku_df)
            cj_final = [r for r in (cj_today + prev_rows) if r['묶음번호'] not in kd_seen]
            all_rows = cj_final + kd_rows

            self._log(f"✔ 대한통운: {len(cj_final)}건 (오늘 {len(cj_today)} + 미발송 {len(prev_rows)})", "ok")
            self._log(f"✔ 경동: {len(kd_rows)}건", "ok")

            if unmatched:
                self._log(f"⚠ 수동수집 미매칭 {len(unmatched)}건", "warn")
                for u in unmatched: self._log(f"   {u}", "warn")

            out_dir = self.output_var.get()
            os.makedirs(out_dir, exist_ok=True)
            out = os.path.join(out_dir, f"2026_플레이오토_송신_{today}.xlsx")
            save_playauto(all_rows, out)
            self._log(f"✔ 저장: {os.path.basename(out)}", "ok")
            self._log("\n완료! 출력 폴더를 확인해 주세요 🐸", "ok")

            if sys.platform == "win32":
                os.startfile(out_dir)

        except Exception as e:
            self._log(f"오류 발생: {e}", "err")
        finally:
            self.run_btn.config(state="normal", text="▶  송장 파일 만들기")

if __name__ == "__main__":
    app = App()
    app.mainloop()
