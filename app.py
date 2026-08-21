"""
ECNG Supply — Mobile Snapshot v2
Same styling/branding/data as the current app, restructured behind a
home screen with three sections: Outstanding, Transacted, Pricing.
"""
import json
from datetime import datetime, date
from pathlib import Path

import streamlit as st

import excel_parser
import onedrive_source

st.set_page_config(page_title="ECNG Supply", page_icon="⚡", layout="centered")

NAVY = "#002F6C"
GOLD = "#FFCD00"
GREEN = "#43B02A"
GRAY = "#989898"
DARK_GRAY = "#333333"
RED = "#D6483F"

STATUS_COLORS = {"Active": NAVY, "In the Money": GREEN, "Expired": GRAY, "Transacted": GOLD}

LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAIAAAAB/CAYAAAAn+soHAAA48klEQVR4nO19d3hVxfb2u2b2PiW9hxZ6DaEJIgoYQEVsKGhiuYoV7F2v3ZPjT72KimJBQew9B1GUa0UlIFKUKoSShBJaEtLLaXvvWd8fJ2ADBVK9fu/zWJKcs2ftmTWz+hrC3wQuF4uFC7NETo7bAsAAYJNAwOTIZ95dkpazfOPgmuq69O2l1T11Qf027SgGAGrwwMyQmuS0Lm1AUv7csW10XriQXx8/JDXnxotO2G3XqCZoHfg0pbtcclRWlnITqQaP3Qxo+AQ1JZgpI9MjPJ7MA1PstEnM/HBp2jc/5p+am797xL6KugHVtb7O/iCT1x8E2AIMA7BpjUtLwAB0GyAlwu06wuwimJwQvbtL+4S1cdFh866cOGLFaSf0yK3zG/u/ITIysik7O0MRETcuMY2HVskALpdLuN25BHgsAAh32jD1zZz+n3yz+sydRRXjy6rrhlb5mPxeH2CZABRAZEGACAQQCebGnXNBBGZWDDAUA8wSTIBmg9NpR3KMjVOS41f069nhk3NPSvvvacN7rd1/MqSnu7RRo6DcbnerOxVaFQNkZGdLT6YHgMcSACzm6Nue/GjK8p8Lz926o3hIhc+SAZ8fUAYgyCQhCAQCQzT3FiMCA2AwmJUCTCUhdYRFhCMhUqp2bRI+79W17ezXs85fQES19W8oXa5Ubk2M0CoYILTj3QyABYBZH6/oNe/bNbeu3rjztPIao6O31guwAQhhkiABAjG3Dtr3I3TMM7OlFCylQXci3KmjfXxY4cDUjm9N+/e5b3VMjN4cWvnWwwgtOonMTDQqSyLHbTptAm9+unLstHcWTsrfXjShvM4KswJ+gJRFUhDQ+hb9UCACmGFBKYIi4QiPQIfEcG+vrklvZY4bNO3yM47bogAg3aXxwiyrJXWEFpvQjIxs6fFkWgJA1uzPey1Ympe1bXflBbuKKsGmv/6IJ/l3WfRDgQjMlrJgsiadYUiKcfiG9uv8xn1XnPTM0AFdNwMAMrIlfqXoNit9zT2gy8XC7SYAUMycdM4ts91L12ydVFLpC0PQp2DTmQji777wvwcRmBUrWEpqjnB0TIr09uic/MAXM66eTUTVLSUWmnmSMyTgsZx2iUvuf+/aL5fk3rOntC7F8NUCmrSISDa29t7aUM8IFkylOcIj0btzfOHJw/vc9syt4z80+ZeTsdnoaY5B9it5EuBX5i3tP/2dRY/m7a46o7aqCiA2SYq//VF/pDggGhRpMTGR6Ner/Zsv3Xvaw327d88D0jUg54DDq0npaOoBQqZdpuW0a8i8680bv1+1dVrBjhINMC2S4n/uqD9SEEGxYbF0RsiUpIiSEwZ0vjb78UvnmurAxmlSkdCkkz9lykx91qyrjXKvt9N5N7/20vJ1hePqqqsYNk0BkE059t8JRAS2LBMWadFxURic2u7Vb2becB0RBZCRIeHxNJlIaCoGICBDAB5rRvb3459/7/tnN20v6aQMn0maJpn5n7PrGSBBOBzdhgjMpsVCc4q07klrb7xo2IWTJ47amO76TlvkHm02hTxo9IXY78YNs8+1Ls167+Z5X6+dtqekSkAoE0Jq+B9X8vZDEEEFDKT16RDYV1nHxcWVDtLkXwt1AkixyRZpPbq19U44Oe2SqTedPXe/At3YdDYqA+yXWZKAM2+eNX3Jur03le7Zo8ihgxmiMcdq7SDAZLJp6cekvLKzuKrv1p0Vw4hMxaDDmgcCLDZZRkY61Bmj0+75+Mkrp/oNJZm5UYNLjbYoLmbhdrsVM8f0nPiflz77Pu+m0qIigxw2+qctPgCwYrI5bOjVOWl+XJRzA2k6gMNfOAYk6cQ1NXX4aun2x0ddPeMZXcIaNSpLulyuRpvPRomZulws3ESKmWNPu2Hm11sLKwYbgTqTbJr+v27XHwwhVzCLCDsFLz7z2CVL12yLEVJcaZnMoMM/dJlBpGsoLykxvvsxcHOnMx+inHkP3pyTExoGjWAmNpiTMrKzpdsdWvyRVzz71ReLNw0O+GoNSKn9ExcfAMBQII0inbLw5KG99vXv1q4gwqEBiulIhS4ziGxSD3hrzO2FlTcdP+npt5k5AcgQjXESNOgBLpdLeDIzLWaOHXnlc18uWbV9CHHQJE3q/xRl72BgZgVNR0JM1OKAoWjs8X1zJVklEEISH/muZQZIkmaa/uDqgrJ/jZn8/MsCHsu9EAJomEV11AzgcrmEOzeXmDnm1Ote/OqHn7YeCw6Y/E/e+fuhGE6nDd07JC4iIq7zBZVlKW6Izs0MQAqbv6rc+O6nred0Ot31jMhxmzhMpfJQOOovu925RB6PNfD8x579buWOIUoFjX+SmXco1CeKSKfOvmEDOi4FgG/WbD7GZJEMZVkN4wKANKGTFbT2lvpvHj352acdNmkNnjJFP9pHHjEDEIDBU2bqDvtH1inXvDBtfUHpJYa31oAm/5EK3x/AUJA2sutizS0Xj9kCgLbtLB3hCypANNx8C50EJP111cEft5Tect7tr9y/ctYsA4OPjgmOmAEmT5mpr5x1tXFF1lt3/rRl362W32uSJv7RMv/XYGYWmsbtEqKXCCLWpODyat9IZRogOgIT4E/HAEiXenVpqTn3uw0PXOV++2xaOcvIyMg+Yvf6ETFARka2nDXramPaW1+f8f4XPz9cvrfYJF37nw/hHhGUEk6HRr07Jy9mAEt+KIyvrfMPhGUC1Hj+EGaQ0IX01tTZvv1x64dvf7Kkt8eTaR2pZXDYH3a5XMLjyVQ//5yX8tKcFa+Wl9XYyC7EP8qv/xcggCE0oZNZMu7E3gsB4L7ZH6bVBTkerLixI5+KQaQJa+u2Yvnw69/NZWZ7vWJ+2OMcHgMwk3shhN2m8Y3PfP7Glu2lSSSUydwwDfR/EApCol1i7Oorxp9QTQCCFp8TCCqGoCaJ6DEgScLatK28T/qVz7woPB6LRmUdtig4rAXMyPQIynGbZ9zw0s1L1+8eDStgshCNXHnx9wczQ+o6dA0fGqYixazvKa46URkG1aNpBiaSbPjN1ZtLLr9xavZlyHGbhysK/vJDLhcLjyfT+iRn/dBVm3Y/E6iutkj7/+be71Gf5SiiHJo17sSBywDw/TO+bFNaVZsKDsl/NowmSe5gAKQJWVNeZX3xQ96M/PydPdxuNx8OE/zVB8idm0kOm8Tjr341a3thKcimHVZs+38NhAM2/qFgQepks2HtYzecvgkAFv64YWRtEHYQmWwqGtivm7DpGppi84TiBsSbd5Q5L856/0lJ4Nzcvn955PwpA6SnuyQ8Huu8O1+7fvnGogEgy2TwPzKThxWDTXVIXz4zM2k29EpJWkpEBjOLoKEyzaBJYAVHmAOd28a8YAYNXwOdd4cGkYag31xXUDp+8v+9c67Hk2n9lWl4SEpcLpfIyYFatGpL4sKf8tyG16tIin+c0kcEwLTQLjna7NGlTRABkw8qyhVThFPDgJ7t5wNAWVlZxO7iyuNhGSCSWoTTtlsyvS41aWsq8ckASBfCW13Li1fteJ6ZYz2eDfxnVsEhFzQ3ty8JcqtHXv7qiT37vPEkWP0T4/pgskizI9xh+y5j7KAse2QUwVK/0egJUCAhIx1i94OTRy8BgEvuzx5SVhOMB3GQNRs0gS+euf2cori4KAnTOpKo8JGRyxAQbG3eVdkm/YpnrgPcatSfWAUHX9CQzW9NffWL3is27L5IBXwKgv6RWj+zYtJ1OO22NY/ccObTndtFVbIFIQjql8+wIt3OcVHhnyUlJtYQASWV1ecZJkswlMNuQ+fk+K86dIj3O3Q9GNIkms59QlJIs7ZGFeypuPebFRt75eS4D+kgOuSOZmb6enneyxWVXp20Rq+2/vuAQXZNcLdO8T8Skf/4vh2ec0RFkbLULxo9s7DbJHVLSfiQASjFztJK7+kqGACY7VEOBG/810mrBVHpruLSLXDY60uKm4jkkINIFZV6wx6d/fV9f6YQ/oEB0l0uDW63evPjFWNX55eOgBm0QP/MFG4iMBTL6HBpTjr7hGVgptcevnR6uwTnPihIIjABDJIiwk57Hpl8whIAdPVD759QXhvsBLYMSBtFhtmWXnrOkDwAsGnyKDICjop4afnr1Nq8ovOf++D7Xh5PpjqYQvgHBshx5zIz0ztfrnigpLiCQ2ZfExIaenjrPF8YCpoNuiZ+OOfEtN3om6kTUVnb+IgZWlgEsWKLwRZpNrRJjP5v/35ptYLA6wv2TqzzmYAgS7fbkRgd/plpKUgpEGa3N8vbMkCQUpVVBWzv/nfZvQSwp+SFP5wCv2GAkJzwWFmzPu3108Y9Q6FMRhObfUJKgOnwEuebGczMQtc5JtzxNRGp1L4ZAED3Tz79vYRouw+mElAgh10itWubDxWDrKLiiB17S89UwQAACCEIjjC7GfqBEBFuD53RzUA/CdLY71Xbi2rOz1m1JRU5OZbLxb9Z89/84M7tS5okLFy+7e6K6oBOUqimWhYiMEyFlKSYfWOGdt8Oi0iAW7xhwm+glAyzS+rfp/0SAHDGVnB6ukueOTJ1c6fkyEWwOQSYZdu4sL0fPH75EgA87sG5Q8przY6ApQAI01LYsbcstN6hFhLNRj4zA5pQe8t99qmvfTsFAC9cmHVwBnC5XAKeTGt7SV3KzpKqC9nv5SbV/BkEQSiv8Uadnd7/wg7tYvMUa4JaCROEvH6SIhyy+O2HL1sFACtnXW3k5LhNi5m6tkt8yeGQgGbnCKf9TSFErRQChcX7Jvn8JkBCAaFFqKzxtdx7CJKWt47XbdmdwcyxOTlu89d+gQMMsHAhBAG444n3L95ZUmuDJlVTcmvoyWyZ0O1fLNkYO+2u8Re3bRMb5KCl/sLl+uePbCz6FCzoNoTb9e8kUTUAXPPI+xfc+Eh2KgB+97ZxX0fY5Z7wcDuddnzvd5gZK7dsSarxGuciGGASvyjOogX9Z8wgkFLldardrU/MvZgBGpW18Bfa9v9PTo7bUsz6pq3FlxleH0g2A9FEHAgYqKj1npR58rErBneLmxTXJkljQ6kjjpxx4+TJ//I8BZtdR1r3tosVQIIIy9Zuuyd3W9FFAKC1b1sX6XQsToy0bXvstnM2AaBr3HNPLq70R0FAtaaqZ5IStdW1vGDZpom6AOe4Fx44ZQUQyvQBwPdM/3Tg9r1VPcAmMzeDz5+IlGWhxucfwC6X+OrlWz4YObDDY9LukGwq43B5gMDcITlmnyCixuCB0ENIOiTXDezbdj4AXrdjR9y2PRU9NxbsHahLgqVAPTonfzKod/vniMjQNcGlFd7rjEAQJH5LeNNZ/IcNASuIooq64a/MXTEIcKuM7JBJKADAU7KBACBn5ZYJ1T6LIEWzdKggAsEMoqbWn4qsrLCgqWje01PuGXFMpw/1sEidTWX82T4iAGBmm81mXTVx2ANx0eFVsEI2RUPoYoQSOyPD9J+yppy1EwCefX3ROK8hHBbzyIWLVicB4NceuPbjWfdNnA1muu7R9wcWldcdC2Uy/6r0nQhwOo46abdRwABBkFXtVfrXK3IzAaDkhdCah875HLfFzPre0qqzVTAAombK9GEQwFztNRKfev3LDgB48OAp+sKXb7pqYI+EJVJ36KSUecjvU6gGJyk23PvMnefNTO2a7NEc4QTFDWNgZpa6hv492y0WREwAfli74xTDH0BtwIp8bt6qAQDQvj35EhMT60DEC5ZtmeQzyAaC9avnQAqBtolRDACKGVW1PkA0kEOPAiSFCNbWYuO24vHMrNe33IXYf/w/NOPTgZW1Ri8oQ6ERi0b/DPU58pahhL7s5139AMDXtS0RUeXT1w47LSEu7AdIhwY+OBNwvaMm3GlfWFHjo8mZw59MiLHXsMXiaGUBAQylZJRTclq3lI8YgGJ27C6pPBEqCL/BVFJeO5oApGa4dLhc+PLLNUklFXWXWL46/v3mEUKgQ2K0FqKXETAOzc9NCoYAKd66q7zrY68vbA+AXS6XEFu3LhAAaE1e8fjaAEuQUM3KnYJQ5zewt7T6WADILYHKyMiWI0aMqDvjmM7j4qLtP0BzalD8x5ljZs2ms10T84mILzlt6ObEaPsjwhkumNVRngKkIHSKcGgbpt561s8A6KSrnz3eZ1BXAIZlWthZXD6QCMhdXU5wu9WbC9deU+3jBBBb/OvNw6H+kW3iI6oAQDHgD5gtoh7Wx5+suiAcK9bljwdA8+fvlWLlyraWrgnetH3vCDMQCKm7zQgiECwLe8sq+zGzQFJfru+SpV59YmfdReltTouPtP1AukMD/4YJGIqlU2MakNZ+OQAgI0Mumn39zM7JEaVQQhJwxOoXs2Jps6NtUvRHgsggAhfuqZgYMFRIuTMN+INmf0txJPLjDGa2LfxpywUBbx3/wd6TQvp9Xrz62aolALDkpy3dU9rE9kQgAKIWsA2FQDBgYNO2vScSwCsjNrMA3CpoWNEVNYH+sIxGzV0/PBDBMuH1Gf0B6PBkWjc+9mGHGe8vPhlwq+fc5bVXnN5tXFykbSn9ShzUB2Eo3KFtfyPr0jwAlIFUGRsbWzmsX6f7bGFhxEfoxyQCoJSMdAp12nFpHzMAtXWbo6LaexYbQQAkIIDSirqYi+96ORpwq0vueT2jvNbqA2X9UXQyw65J7pMSBwD4blVemGmxo/6PDZu2owARJAw/giafqJjjkJNjCgC4+qEPjqkLWHGAanb7NaQHKK7xmXEX3vVGdwCorbICr3+y9NUnX/viHIJbPXHXzrorTj/u1Lgo2w+QTg0MgwEmTUdsZNgySeQDMoTH4w4iPV3zPHnFrJ4p0avAQhIdUTq2AmnUPiFyY9YN49YAwFlPfDm8LsCdwKYCsYBSltQdYUWVvjSnTeLngqL7fTV1qG9newBEBJgW4uOiaNIZQ3UA+GzxOquyxgtI0VI5tQQwV9QEE++a9lF3INTKBrtKykb6ggwI0RIWKwGsTBYOwzJHAsBbj2Xu27G3YvMr83/+6IP5ywcCbvXET+94LxqVelp8lG0paXYdShmaTUfH5LgcBeDG6Zdrnq9XDEJOjmmYisaP6jc5ISEKbJg4XKcSK1bS7oDDLmdSqOEFVdZ5r/MHLIYQocIOQfAFLcrPL3I8+uo3p20v8fYGmxb+oDgzQ2rwBwJF4085bhcA5O6oIn/ABEQLKAGoD7wKoWp8Jq8rKBoMAMImAV/AHGn6/aDfezCaC0RsGBZyC/b2IyIYiqlz25icjVv3YfJjc98866ZnkuHxWM+582svGtVmXHyUvgxkc0Q5yRx1XLelAGAUlkb/55WFH742Z0lvAHjilrNXjRzY8WXdGSGhrL88BULHP8u4KN0/ZeKIeQCwYMGKuLzC0tEwA0T7F5gIlmEitVeHsz74cuXdVeXVTLr2hwOdOcQAwYCxr1e7qH0AcPrxPU50OMOAPzNtmxgkCGYwSEqpk+2agPAbbMvdWhRf71VpGaKIhDKC8PqCx9l0ASJiTbd979AtVNX4+y1ZU/SNa8aHSft1gjtPGzY2MtK+xi5Fyd1XnJoLAGaEPnTzHl+XWfOWunVJbFgnanOfnnJnz06x29kiIvoLhZBhQdooOTpswZSMEbsA0CsLci+o8FqxIJj7y7r3y9HcbcUTNhTsGQE2CAfxmhIAWAqd2sb5FIdCsMt/LtQNq+Wj3syMvB372ltKQXy/cmsbmy77wAi2gAJYD4KAZcJvmH3f+XxFJwC4NuPEbVHhuh9m0KyoCfR967Pcb9+dtyAZcKu77jqnpl+vuDMG9e5wLxEZALC+YNfousoKlbu9LOOmx+aeBeSYRFR19og+18QmxAg2rD/1ELKlEB7hRJ/ubV4mImXTNV63Zde/gv7fWkbMADSBHXvKY2tq/QJSHPShzFBks6OwqGIxESkioC5oxpmmiSbLCD0cEATMIExl9dhaUtdevPnpig6llT4J2XJ5fyHZBK6sNRzPv/99ZwC4auJxO8Ic9g2QNg0qGNy6q7zvg7OXLnh33tJkAPjh1bv2fPbCdW8ALsHMcl9F3UhAiaqqWni+Wf0fZnZi8BT98dsmfDmkd/Js6QiXh/IQEkGBpIiP0Aqyp175NQC67N43eu8oqh50SMcY4S8bPglBGNwnJZQZwkBSTNgIKIXDVkqaAgyACNU1/ij3jE+iRESUbahuc+hgVmjJCySILJMJ1dXedAIQNCxEhjuWkK6DQQIqaObvLEt7cPaCBR9+uSYJALqNm24H3OqVeUuSK2v9qWALBGXuLfP3HX3V9IfEylmGxenaVy/deGu3DjE72II8mChgSynd6aTBfTq+J4h8UhDvKK74tz/IDqJQOf7BKP7zF2JBUCgsrVwEACaztn13ua3hE9UwhN6FLRa67No2dpR489MVRq3PD7RwzQcRkWUYqPUaIxUzKWYkxUYssukCUIpApIWYoDztruc//WbevAXJ+V/cHGBm8nyxbkhtQIUBsCBIM7y15pq84ttum/r+cCDHIqLa80/tf21icjxxwLJ+twEZTDImXFbcdPHI6Qym/IKtMevy9463Aj6GOPK4yH51yqGLYPrAHmUAsGF7SUJslPMYGAFQS4naAwQSAqZF7325UokRA7uOU4aF+phHS0LANFDr8w/eubMqFgCG9eqwwiZRd6AREokQExSWpd3+yrKvP/wwpy0RcVFZ5ZiAoQihzEIiTYjKijoxZ2HubGbWMHiK/ugN4z8f1CvxYUd0jM6m9UuomdmSjjCKjwqbMWZI71IC8b2zFl1TVmPGI+RKOuJ54fpMYQlVNa5v5w0AcNvUOfqefdWiBX0AB0ACFPAH0S4h+lhRVuGNRqNE0f8CSjGUMuvduSYBFgALDIsARQQQ2KwLcPTTngUDAODh28/ZE26TWyA0OpAqRqQRG2b+jtJ+D3uWf8XMbYrKqofANA+0YGFAEFnmjr21vYdf/vSLYuUsw1KD9S9n3PBQWtfYb0lz6FBsEhHDYoqLtNXdd/mpbzFcQjHLnB+3XBL0+Rji6MxiAjOYEBvt3JmRcTwDgC5okGZ32utNwJa9q6n+v5Fhjt5i1cZCf3NopXaHg+zhURrpYRqETWPSJYRNQrNJViRU0CIGy1q/Et8szUsDACKyIsId3yNUjf5LJQ6RBmVYazbvTUub+Mg3lTX+niFe+uVoZSKNzYC1dkvJlRfd/crlUqw0iEgtmn3DFf17tdnFltDAVlA4I2RijPPpS84Zuhlwq1FXTZ+4ryqYCj6Ia/cwwQwFXYfDbltHRH4ACBhqkC9gUks5gQ6GTduLA8Ju15JgWY3WwOj3IECBBYb1S1k1ZlD7a8M08+YB3ROeGz6wy1cJUfpXgs2v2iVF5nXt0qYuMiqsCmzVllXVDdpPTKe20Tl2/SDFFIIkWwG1Ib84NWBYiQfiXb8eW5Koq6xUi9ftevXhmfMHA7DCwsJ2nH9S35M6dUosZEOzt09wFL/iHv88XC7BzDGlld7phtfLDQqKMUPTJGIiHcv3/2pncUXvVpX5ToSCXftYS+vRYcDi79eDwh3URAQS2OLVm4t63XHpqMWfv3jjhvX107J/S+9mdm4oqYt+bMZc9fUPBdi9Mt+/v/P4KUO6/rgyd3cgwHAQfqeRkxD0i5z+w4KF2qxqvGNbsXpt/tqPt27ddXLXy18uuHfKGVtcz887aU5O3tcpCc4Hju/fvxgAzqtuf3X+7uq2EGwyGpARzSCnTajj+nfJXQaAmfUOY+/vBdMAUfNfcnlQMENZ7NT8/sBRy7rDGgcgkqSqq+rC3/t8Zc6aNXtTBw68qRLpqQo5uQx4LCLyAfhN7rTbDSA1w3Z+z8iSGRGOLZU1gf6hiDp+p8L/+THNDEE6WVsKyzucfvsbb/LCrOFE0Nw3nJ3PXDEIiKmmF0vEggXjEi94ZP6tgdoaRZqQR7sXCGAmIe0al545etDa6QCemrUwusYb7BKivyW9QCEIIYTyenHM0F7DtDp/kNA4uZSHBDMESVgbt5bG3/pc9ps2bc64IB7UAI916V2vdV6zvaLn+JG9rTJv7YbC/IrqT2dOIQBBIgp2He3BiMumfbqr1N9fGV6Fo0hXYyIJy29tKtg3NHXCo3OZsyYQ5Uqi2EogXdNEjnnX7LAZ5VX+ZBKwGlIGH8on1GWY3fbTyYO7VgPAkg35g0wlokOlxtRq6iyJSDRbX18GJKmguXTDnlMz73rt/ncfuezhB1wukdync8QPG3d//sIH3wuv31fJrPwppz5IEU69lgbe8GPb+Ciu8Qa6CphQDblniEgSG2b+7qqz+p73n4+YsycQEQTlmBNun3nF/MVbJ6qgz4IUDVugUD4hx0U7F4nQxQ5UUVd34n4ztUHPbmQwKxYOm87NZphKIf01tdaXP+S5r334jSFut1tdd8Ho9celtTsnYCn463zRARNtdu2pSN5UUNKNhf2CPRWBC9dsKT7ONIINdqGyIM3w15obt5ae1fWMh//LzCm3PT632/erd8/w1dRapDWCN4xZOmxEvTq1+Z4BOGwaV9b6T7SCwSZTtI8KzLDZbCQiI8Io5J9uljGJNNC+fdX02dLtrzNzNJCuvf/4lZ/265rwlC0immAZQdLApEHB9Jsw/SbYUI0iOkPxcA2W3ywsqjo17dxHFr2/YNXXe4sq7KRTg+8mJgKDBUU5tX3vT718HQD4Aoa9cHd5l3oZ2yo6rDAzw2bHjj2lBWJfWXUJND30y+YYHCSILGtXcV3fUVc+O91uW2Ka1ona0rduv79jUvgSknYbuL4dDZEGIq2hLdH/QAMJzTJ91oYtRZ13FVV1QagKrsFjMMOCzc42TX4lBVUBwPWPzhkRZEqBMg4VU2h2hBjAhh07SwvE5m1Fa2F3NBsDAACk1Axvtbm2oPTSi+957SpBOSYRBe+adPJlifHhpaxIEDUxPUSSJBQRq0bTzJUim02jQb1SvlccEljr83eP8gcZoOYptjlsKIUenZN00b9ne625C3KZGWTTZWXJPuvLZXkzH3/50zQAanLmCfnnjUm9PjI6kjjIZlNLTA5FaRvldKFQUElE2KhudHr/+QBYElC4t2yEaQSbPdv6cNCjc7IuenVO1purYcGvwcwk7Brt2llGs+evm1dTVJMMQM6476Lsscd1fS4sNu4vS8NaGRQ0naLCbctuzxy+CwCyXvy8bVm17xhYof4QLUveL/hVdssOsWFr0UIIAebGu4vucKFCHQisvMKKrsOue3YmMxNSM2xznrzqts5JjmxIu04HKwhphWClWLfb0SY2co5Zn4L4w8/bTw0qLQqhViuthpWZwZpNx08bdy4WfbolFoeHO1qshJWF0FTQa2wt8p595o0v/ge5niAR8Ya59105oFfyVlZCa3WdQ36H+uNfhtvYe8Gpwz4DAE0K7CqqzAj4A2hKT+tRQTHC7DquOOcEXfTvkbIPSlloKR8lM0gTuq+q3MxZteOOm//zwRUALCLyXjB6wElt28buVRZRa+kccggoSBtSkqNX3DxpxE4A+HptQZuSitoRMIPc4gkgvwVDCOn3eYMrc3cuFMlJMSuiwySDWaARryQ9IooYIE2K2po69c5Xa6e7nv5ooMvlwj3XnrbddeWYqzt1TAKbqiWKag8LzIp0m0bx0ZEv1V/rSs++suC0Sq8VBWKrVTWLqC9YaZcYhadunegVk8cfV61r2m6QBFowV4UBQYJRWloT8X5O7idTpkxxAKBrLhj16eghKeeHR0awClpWQ2v/GxuhXoIkEqL0cs8dF34GgCSBc7fuPdfvC4Ba2/EfutQARJTXr2fyHqELKouNcuZDaEDztDA8NG0gQYKtzQXFKWOuf+1jZnZ2Sr/U8fpDl3uO7dP+6ujERI0DptmaPKpQrITNgd5dkz9p2zepBgBmfLC4Y3F53YmwgqF859YFhpBIiA7frEvyC0MxxUSF/RhywDXtCVB/PO5PAWMcJAbJgCSY5va9daeMveb513fmvOHvlH6p48f3/v3KZaf1nxGbmKBzwDSoEa5gayio3qXhtItAx4Sox03FAgB/nLP+X16DIgFu2D2BTQBWzLrDzolxkatMBQgi4oG9O6x32Ovvtm3KwYMWQdgksxBsKgKIYFomFIdyBev/YSEQqKvyLVhecF5qxiN378x5w1/nD8pZD55//cVnDpwRmxCvsz/IIS9ei0JB6KJHSuyP7z1x1SaEvKyOgp0lVwRra9Eq2+uzEg6dqEfn5FVA/fF09vDUJZFO4QezbJrTlVkKiXGj+panJDg/TooN29ymTaxXCvii4+M03RmuQXNo0J0aNKcGadfI7nQyBG0vNv9zwmXTnmRm6Qteqc+8N/P6EwalXNa+fRKxJQWh6T2Gh3wrpRAW7kD/7u0eDZoWAcDVD79/4t4yX3fAanXt9UOnrhCRTln10JQJK4AQA9CZYwZsbxMbvr1eD2j0XUUgKNNQ6/OLtZNO6POWtebZ3v++dkL30cd2633R2H6Z7ROc58c46dJxx3V+s1/XuLeTo+1vh9vwdmykfJt9FW8GTNXp7S9WxjPPNIPDH9D+++zVb1wzcdjJXVPiN0E4NTZb5CCwIGwi2iGWvvHIpV8i3SUJ4CVrtl5T4w0CQrb06fRHMBSEhoTosPVJCVoZ4BKEwVN0/mmmeeLl0576ft2eW9n0m6EIXBPAsBCTkIBTjus688MnLr/mUDN0iIZ/+3fTga85deDU62dNz1m59ZqKimqNpGi+fDullD0sQgzp1fb0JW/e+jngEq4Zx3ae9s7iNTVVteEkiVqb/IdiUzrDtXNG9Hjiw2mT7xo8eIom0s+8kImIh6SmfBUV4QAs1WREk02isrzUmvfdhqv7nz/1o/XbitsAANIvdQwePEVHerqGdJf2+0WsJ0g5dFLM7Jjx4ZKkyx5495Rx18y8o7CkWoTZtVCuYDNNN9Xv/thwufz7N275YvDgKbogt1q8uuBur19FkkDDrgpvAtR3PxHRTk0d27fzJwC4a9eTFTEzEREzc3z3Mx/aVLC9OIF00ei3XP5CCIEty5S2MK1dQnje2aP73/rSvef+11RAhivb5nFnBh+c8em1a/JLTzV8dZV2h64VlVaLwr0VLEgkhYU5++4uKnPodkdsrc+AaSog6AN00XzV7ZaynJGRctywLqd99PQ1XwAusWn7NZ3GXjszt3BHsY1sssHJJY0NAhRbJLp3is/Pm+/qS0QGg6GFTLMMKYjKTrzimW8K9lRlsgpaAJpEDDAzIIVmGT5r555Ajw8+Wzn/rJtnvzh32pX3EFFVaobLpgLWt7v3lp1RUOS/tLK0FNAIJCTYNAGzApAE1PkUhFAgAulCay4XFgEWS5vsnRKz4vMXbvgCgwfrYpXbuOeZdvfsKal1kCZM5qaZu4aAmZVwOKldYvSnRBREukujHAr1CEp3pRIDGNa/qyc60kmwrKZVrEMlypIkq31lFeq/izdf2//cx37699Mfnbpljjv48K3nbN708T1nXnxqryk9OyeUSBDYXxsgMoOkkwKBSUpRr6s02+IDIcs+zKkFh/fvdqM/aBJWdlXL1m7uunLz3ktMX52CaD1Zv/tBAKCUiHRKGn5Mzw8BICOp728qaUI1dcyODmMf+Hl3UVU3EqwaK1niT4kjgJUyoYQWGxuJ7inxM2bdedJ9gwYNqgSAb1Zs7OWa8dW9G7cWTyqr8gIqaJEUoiWOWGK2SHfKkQPav5vz6q3/Sk112XJz3cH0K56euXjt7inK8DWdAt0AEEGxAdEvtd22ddn3pBJRoP5PB26P4PR0lxREvn492s3XHE4wN4+ThRkACY00qIqKSrVy457rJt776ZqL73/zMmYWJw3ts3npGzdfeumEYyd0bRfzvSMsUrIJglJmc8YFKNSAltolhFU+duM59yHdpeXmwnxk9pdpq7cUXayCPoVWlPP/G1jMelgYBvRo/54k8qenuyR+n6WalNSXGcDkiUPfjIuyGbBUsyYxMUOQJoWyAta2XSWd5n6z4bVhFz/13UOzPxsEANNunfBxwWeuEy8cN+DWnt3aFDqcEVqoH4Qy/7L/T2PQZzHbwsNFapek248f1GV76j4ITbrVR9+sea6mNhhW75puVYofEHL+MEPERMjq8eP6PaMA5Iz6Zb5+S3BGhhRzPFbaxEfmrysoPQNsWGhIMcZRggAVavIoZdvkaLN7+/hnH75x7JPpQ1L3AgAzx54yZcaVm7cXX19c5e8c9HoBUhYJQU0htgiwWAl5XP+Un9bNuffYLt3OseXmeoLXPvLuxW/MX/uWt6am4QUlTQRimKTZtfRBKXMWv35rhjnxPAmP50CC6m8mKwMZUAzKHDvwtZjocMC0WoSlGRAQJElYau/eUm3JusLbrnroo1WZd7x2IzPbiKhiwcvXPzln+pQhJx/b5c5eXdsUhoVFSlZCQClF4EZzZhCB2VDo0C4+ePuFJ072Be4VTmcsL1u2LOrj73L/z1vnZWoiB3pDQQBYKRHm1AJp3ds/alpMGQf5zO9+fFAwZ1HauY/+tCGveEBIGWy5C6OJwKxYQUHaHE60TYhYOSS13VNznpr8IREFAaCwsjLOPf2LqxYu33zJnn01ab6AAbABCGERoWEKo1KmLSxCO3dU7/vee/yKR7uPm24v+OLmwKDz//PS2vyyq5Xhs1pTvd/vYAKa7N8j8ZsNH953inVexm92P/DH7pZIT4cgIuvY1E7/Fx4VTvzXPRabFKE7b0iSRhwM+Kwde8oHf75067vtT77/h3NvmXXve5+v7tw5Jqb8FdcFUwu/fqjfqSf0OndIaocVMZERANkkm0xgtugoGkcTYEHatR7to77InnrFo6mpLlv+FzcHLnO9cebmnZVXWwFvq138+kuvKS4ukqZkjpxpMZCemvqHjXCIneESzFmi77mP/JSbV9KfhFLcArrAwcEKihlCl3abDclxzprkhKiP2ifHev5zz/icvomJNXYNOP/BOT3y8nc+kL9z39nltWaUEQgAMCFE6EKfvwIRmIOKO3VMqnr74UsGjnz87t3weLioqCZxxJRnNuYXFEWTTmhtEb/9IILFlpAnHd99zYKZNw4hymLA/YdNcFDiMzL6EhGZE8cOfCgmPpq4hXSBg4MEhJCAqQJ+r1m4qzTyxw27Jy1YuuXTMRc8/fOwy6Z/mnbu1Mt6JIbpm+bdN6no20faXTC2303dUuJ/tGuaqQLGX64+EYGDlhmTECtOPrbrpSOHdCvsVJKq23VNjb5m+qz8wvJYkmh14d4DIIANi6Kiw6zj+nWcTESWy3XIjx4CLpfgrCyMnvzs0oXL8ocKjZVqhS9cb+YoKIsAKSB16JpEbIQNmi53BYPBRX27t9le5wuu69Qm/vTispqLl60pIAs4eEoRATCVYQ+P0k/o3+bxxa/ccveQc59yLvPc7rvorleempuTd5u/rtqEEK3O4bMfxGxBOmT6MR0/WvzqzRMnnpct6+9g+AMO+RIZuX2JiKzp7y68ZtOOslVFu/fV3yPc4plYv0G9gichZMh8ZIONYBAl++oEhNYBmn7R4nVFUEYQP2/Zw1ERTlKCwIe4F4UUW9AcelqXuEXfzrr5wc5bKhzLPLf77pj64YTZ81fe6q+pMkmXR91BpKlBBGYLlJToqL7uguG3LHzFJVJTNxyS2kPuaI8n00pPd2l3XDJm9bG928zQwiIkrJbrcn04qPcBSBBJ0iWRYAUVtFSg1gQHVSBo0r59VVCHWnyCYoNFj87xe959fNLlRGTuyHH773xmTo/X/rvqxcqKGpDWMm7ow4bFlrA7xbDUjlMzTxlSmJ4O4Xb/Ufbvx5++SH2omHjfvvBOk55fXbirvCtJxczNdKtYE4Do4MnvRGA2GcltYq0Jw3sOesk9aT2QIZnfaDfoghd+WLOhsANJpbiRS9UbE/U+fxqY1qFg9ft39SHKZMCj8CeB8j99GSLijOxsosTEmsvOOOaBmLhoYqMVXIPYABxy8YPKDAt34uwRfabMemjSemCKzpyN7uOfnL12S1EHkGW27sUncMDi9h0S6dqM4dcSkelyZR808/rX+MsX8mRmWukul/bwjWe/17ND9IvCHq4R/z0KNg8HRAQ2LCMqMVEfc1x396ysf72mGDrzTOuEy57xbNtVNZZNX6tW+gAAlmXq4RGyf7f4h67LGLkgIyNbut30l5v1sGRZSBRkEXOWs+95j67L3binC9moUbpqtCTqbwkx7Y4I7aRhXd6aP/3qy+vTo9RZN82as2B5wQS/t9aEFmpV2lpBYItZk0P7d/xp+Vu3HU+UyeBsdTilfoe1gETELhdARHUTT+p7XnLbGOaApaiFagkbC6zYZGhaSpvIt76ecf2k+nuC+DLXO3O+WbF1gr+uxqTWvvihMLXo2CG++u4LRk0K2fzZfLh1noe9g91ut8rIyJaPXHfO6lOG9bghOjFeY9M88lu+WwuUMkk6tCFpHZZtm//ApKBpETOLi+97K9vz1foJvtpqg3St2droHQ2IAA5aKiIqAheOTbt+4pmDN2ZkZIvDOfoPPOOIR83IkM55c60x1zz/5ueLtlzCVsBkKf54a1IrBjFbrKQcObR73qJXbjyFiHYwszj9hhc9367YPjHgrTFb/+ITYJpmRGyidurQlEfnPDX5vsFTZuorZ11tHNFzjmZsIENoco418IKp367cWDyaDZ8BQS17RfZhgphNZk0b3C8l7837Jp7St2/XHcwsB50/NTt3e/nEoLfaJK11Lz4AQLEJ0rRh/Tu+tfSt2y4fNSqLFi7Mso5ULB/V+e1ysXC7CczcfuAFT3yzduPuHgTT5FaYD/drELPFLOXgfp3y3n7g9LF9+vTZXlVV1WPMNS+/uXrDrmGKDRNStGqZDyB07Zzm1Iaktlm6+r07T7AURH2XhSMm/Ki0eLeblMvlAhHtvCPz2DEpbaPzGbpG3MBr25sSzCYrKQendSyY/cDEsX369Nn+0Avz+wy//IUFqzftHaYQNCGp1S++YLaIdG1o33Z5P759xwWWYnKFEiuPivCjNuP2K4WXnDt615CeHUbHxTrzWNhka/QRELMphE0bcWz3guf/fe7Jg/p02X739E//9fL8lcvWb9zVUVl+C9T69RhiNpWSsm/PtnlzH7vsJCIqdLlAbjp8pe8Pz2woURkZoUjT2de+kLI4d9c35ZW+HhS6pK/l8wdCdr4FaHL44K7rvn/1ptM1ot2Z977xxOc5G26vrKojkmjV7t39IA6ZrKk92uRdf96gk66/cNzOjOxs6ck8eJTvsJ/bKNRlZEt4Mq3zb3shZdG6ou/2llR3IxgmU8t5z+pNJEOLiNLbxzs8u79wZ85fsj7u39M/eyl/e3mGt7aKSZct0SLxiBCqm2ATrGlD+qcUfPnsBaPj4zs0yuIDjZjGvP8kePWjnJRp7yz/Zn3e3h7ggElCNrtGHQrsWFZ0fJI2sHv8Oyveuvnif9333qlfLt00fc++ul5WoM4gTeitXNyH3NSmZZKwaWOO77V+wUvXnE5Ejbb4QCPnse9nAm+5t+MZd782L2d53kClghZJKZuLCSjk3yNHWDiNGdb9//777JRHxt/88tTl63bcVFJWDZAK+fVb+eqHAlSWckbHyr5d49798e3bryIiX70F1mgBuUY//vbf9cPM+klXv/D68p93X1RXWWGRXaOmjh0QYLHJsl37JEwc3SuzS3L8ppfmLn9pV1ndCb6qCkW6xN9C3oMttiBjoiNx9pjU2XOfumpyjTeIxl780FhNgHpCOcyu88irnrt+Q37x87sKiwBdWBAkG1vbJkKoOw/ZqW18eN60O8bfOe+7nwfmrNqRtbekAmDTJNn8ouhIsb90HsKmJSdElk8YnXbny/ef/6oFCA5F5Br9BZrusqj6ZBIBqEdf+fK0dz5b9dqGraXJKujdn1LVOGMTAEsBSvDA/l2+zRjTd9ErH684rbC4dpjprWbomkKryWg+NELJHCZsYZGiW0rcj1efN/yqWy46cV2ojNttNhXrNr0GnJ6uISfH9Hq9Hcff+sa0H9fvPLeqshKQZJKghpd2M3O43WacNWZAeVlFTd6iH/OGB5QQYMMk0Xx3Ih0tiAC2QtXRiQlRGDG420tzn7riFiIKpLtcWo7b3aR+lWYxgfZrrTYNmPTge5ctXJH35NbdlfEq6FP1ptjRyWVm2HUNvbq2CRaVVJol5f4wkBnaTa1c1lMondmCYUlHRBS6to8uOGt0/5ufuvnM/5r8iy7V5HQ09QD7Uf9CAKDKyrwpkx5865ll67ZNLKv0AmxaR5tsGZKbClAM0kSz9DRoKIigOGgydIdMjg0zjknrNO2z56ZMJaJypLs0LMyymqtvc7M7QdLTXVpOjtvUBPDQy1+c+dF3G7M2FhQNrq2qAQTM+ksbj4iu/SkJrfm0p9C/FJuKwVImJseiS7vozy8ePyTrpsz0FcAvZnSz09XcCCmImQLwWMwsM//9Rubm7cWPFeyq7FhXVQVIskgQIVTY+bdGfRayBdMChCajIsPRIyVu9bljBz/xnxtPe6/GG0R6uks7mlBuo9DX3AP+Gr/meGZOuG3avOu+W7rpis2FpZ18vgCgDAVN4wZX+LYA6quaLVhKg+ZAXHQYUpKjVg9J7fjk7KwLPiAiCwC5XEyNbdsfEZ0tNfCv8TtGCJt468sTdhZX37ZtT/kxZVU+wAgABJOkJBC3WmY4UKZmKUBBSmc4EiN1JMdHfXL2yQOz3dec+h7VR+4a053bIJpbmoD9YGaiUVkSOSGzh5m1f0//NH3Vz9unbNm5b2x5rRVTV1MLWAYgyYIQoNBlrC12G3e97qFC/QuYwCSh2xAZZkdKQnhp+7YJ72WePuCD6yaesMSoX+qWkPN/hlbDAAfATBmZHrF/kjQCCsvKOz312uLzlq/JP3tXceVx+6oNW53Pj9BtXBaDhFV/LRshlNbdRE0uwWAoBhhKAZaSEBqRzY4Ipw2xYbIiMTF6wZC+nee8dO953xBRWf1XRUZGNnk8mX9apdMSaH0MUA9mpsxMj/B4PABCXS3CHBqWbdzbe+rMzwcX7imdsLe0emh5tT+l1mAE/EFAWfUnhACYLUgBQug6NKpHSKkM9XM60A6V6MDv6sOvihH6GzNx/YVaIRe2pgNCg8OhIz7Chkib2JYUH/3t4LQOCx+6dtxXcVGRJQcK0NNdmmtUlmpJGf9XaLUM8GswM40alSVzctwW6neQLoGgyY77np6Xlr9zX/q+usDwnXsruigj2Luoos4hdDtq6/yh29AEAUEDMA1AytDPwC8Lv3/BQrsasDsAEUqzIykRFe4A2ER8uL1Ks9k3xEaH/XxMr7ZbThjQbfHFZx27ThAFftnWGTIjIwPZ2Rl/i7qJvwUD/BouF4uFC7NETo6bARyQpTYJBEymlesLu771xYoUb4DSfli1OalX1zYjF/+Up2IjnT1jYyI77Cuv4lpvkKQg1PmCEILgsOtgZkRHODkuOoLydxStjYhwlA3r143yd5Z8f8pxacVRUY4NU8b3L2iXnLDTNBm/EeLpLi19FLAwq2VMuYbg/wFTsh+gfBGVfQAAAABJRU5ErkJggg=="

# ---------------------------------------------------------------- styles
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@600;700&display=swap');

    .stApp {{
        background: linear-gradient(180deg, #eef2f8 0%, #fbfcfe 22%, #ffffff 55%);
    }}
    .block-container {{ padding-top: 1rem; padding-bottom: 3rem; max-width: 480px; }}
    .ecng-header {{
        position: relative; background-color: {NAVY}; padding: 16px 20px 14px 20px;
        border-radius: 8px 8px 0 0; margin-bottom: 0; overflow: hidden;
    }}
    .ecng-header-row {{ display: flex; align-items: center; gap: 12px; position: relative; z-index: 1; }}
    .ecng-logo-badge {{
        width: 40px; height: 40px; border-radius: 50%; background: #fff;
        display: flex; align-items: center; justify-content: center; flex-shrink: 0;
        box-shadow: 0 2px 6px rgba(0,0,0,0.25);
    }}
    .ecng-logo-badge img {{ width: 30px; height: 30px; }}
    .ecng-header h1 {{ color: #fff; margin: 0; font-size: 1.35rem; font-family: 'Poppins', sans-serif; }}
    .ecng-header p {{ color: #b9c9e2; margin: 2px 0 0 0; font-size: 0.8rem; }}
    .ecng-header-divider {{ height: 3px; background-color: {GOLD}; }}
    .ecng-header-divider.rounded {{ border-radius: 0 0 8px 8px; margin-bottom: 18px; }}

    .ecng-ticker {{
        background-color: {NAVY}; border-radius: 0 0 8px 8px; margin-top: 0; margin-bottom: 6px;
        padding: 12px 20px 10px 20px; display: flex; flex-wrap: wrap; gap: 24px;
        font-variant-numeric: tabular-nums;
    }}
    .ecng-ticker .item {{ color: #d9e2ec; font-size: 0.8rem; font-weight: 500; white-space: nowrap; }}
    .ecng-ticker .item b {{ color: {GOLD}; font-weight: 700; margin-left: 6px; }}

    .ecng-timestamp-pill {{
        display: inline-flex; align-items: center; gap: 6px;
        background: #eef2f8; border: 1px solid #dbe3ee; border-radius: 999px;
        padding: 4px 12px; font-size: 0.75rem; color: #5b6472; margin: 12px 0 16px 0;
    }}

    /* Home nav buttons — shared base: big, gradient fill, shadow, tappable, pill-shaped */
    div[data-testid="stVerticalBlock"] .stButton > button {{
        width: 100%; padding: 24px 20px; font-size: 1.08rem; font-weight: 700;
        font-family: 'Poppins', sans-serif;
        border-radius: 40px; border: none; text-align: left; color: #fff;
        margin-bottom: 10px; box-shadow: 0 6px 14px rgba(0,0,0,0.18);
        transition: transform 0.08s ease, box-shadow 0.08s ease;
    }}
    div[data-testid="stVerticalBlock"] .stButton > button:hover {{
        transform: translateY(-1px); box-shadow: 0 8px 18px rgba(0,0,0,0.22); color: #fff;
    }}
    div[data-testid="stVerticalBlock"] .stButton > button:active {{ transform: translateY(0px); }}

    /* Per-button color identity, matched to the same meanings used elsewhere in the app */
    .st-key-nav_outstanding_wrap button {{ background: linear-gradient(135deg, #003d8f 0%, {NAVY} 100%); }}
    .st-key-nav_transacted_wrap button {{ background: linear-gradient(135deg, #ffdb4d 0%, {GOLD} 100%); color: {NAVY} !important; }}
    .st-key-nav_pricing_wrap button {{ background: linear-gradient(135deg, #5ecb4a 0%, {GREEN} 100%); }}

    /* Back button — deliberately small and high-contrast, distinct from the big primary nav buttons */
    .st-key-back_home_wrap button {{
        width: auto !important; padding: 7px 18px !important; font-size: 0.82rem !important;
        font-family: inherit !important; border-radius: 999px !important;
        background: #fff !important; color: {NAVY} !important; border: 2px solid {NAVY} !important;
        box-shadow: none !important; margin-bottom: 14px !important;
    }}
    .st-key-back_home_wrap button:hover {{
        background: {NAVY} !important; color: #fff !important;
    }}

    .nav-caption {{ font-size: 0.78rem; color: #6b7280; margin: -4px 0 16px 4px; }}

    .ecng-card {{
        border: 1px solid #eceff2; border-left: var(--accent-width, 4px) solid var(--accent, {GRAY});
        border-radius: 14px; padding: 12px 16px; margin-bottom: 10px; background: #fff;
        box-shadow: 0 3px 10px rgba(0,0,0,0.06);
    }}
    .ecng-card .client {{ font-weight: 700; font-size: 0.98rem; color: {DARK_GRAY}; }}
    .ecng-card .meta {{ font-size: 0.8rem; color: #6b7280; margin-top: 1px; }}
    .ecng-card .row {{
        display: flex; justify-content: space-between; align-items: center;
        margin-top: 6px; font-size: 0.85rem; font-variant-numeric: tabular-nums;
    }}
    .ecng-card .row + .row {{
        border-top: 1px solid #f0f2f5; padding-top: 7px; margin-top: 8px;
    }}
    .ecng-label {{
        display: block; font-size: 0.65rem; text-transform: uppercase;
        letter-spacing: 0.07em; color: #9aa1ab; font-weight: 500; margin-bottom: 1px;
    }}
    .ecng-value {{ font-size: 0.92rem; color: {DARK_GRAY}; font-weight: 600; }}

    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {{
        border-color: {NAVY} !important; border-radius: 14px !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    }}
    div[data-testid="stSelectbox"] label {{ color: {NAVY} !important; font-weight: 600; font-size: 0.8rem !important; }}

    .ecng-avatar {{
        display: inline-flex; align-items: center; justify-content: center;
        width: 20px; height: 20px; border-radius: 50%;
        background: linear-gradient(135deg, #003d8f 0%, {NAVY} 100%); color: {GOLD};
        font-size: 0.6rem; font-weight: 700; margin-right: 5px; vertical-align: middle;
    }}

    .ecng-price-card {{
        border: 1px solid #eceff2; border-radius: 12px; padding: 10px 14px;
        margin-bottom: 6px; background: #fafbfc; box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }}
    .ecng-price-card .hub {{ font-size: 0.78rem; color: #6b7280; }}
    .ecng-price-card .price {{ font-size: 1.05rem; font-weight: 700; color: {NAVY}; }}
    .ecng-period-badge {{
        background-color: {NAVY} !important; display: flex; align-items: center;
        justify-content: center; min-width: 92px;
    }}
    .ecng-period-badge .price {{ color: #fff; font-size: 0.82rem; font-weight: 700; }}
    .ecng-curve-row {{ display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 8px; }}
    .ecng-curve-row .ecng-price-card {{ flex: 1 1 90px; margin-bottom: 0; }}
    .ecng-section-pill {{
        background-color: {NAVY}; color: #fff; font-weight: 700; font-size: 0.82rem;
        font-family: 'Poppins', sans-serif;
        padding: 6px 14px; border-radius: 999px; display: inline-block; margin-bottom: 8px;
        box-shadow: 0 3px 8px rgba(0,47,108,0.25);
    }}
    .ecng-page-title-divider {{
        border: none; height: 2px; background-color: {GOLD}; opacity: 0.8;
        margin: 4px 0 16px 0; width: 60px;
    }}

    .ecng-empty {{
        background-color: #f8f9fb; border: 1px solid #e6e9ee; border-left: 4px solid {GRAY};
        border-radius: 12px; padding: 16px 16px; color: {DARK_GRAY}; font-size: 0.85rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }}
    .ecng-stale {{
        background-color: #fff8e1; border: 1px solid {GOLD}; border-radius: 6px;
        padding: 8px 12px; font-size: 0.78rem; color: {DARK_GRAY}; margin-bottom: 10px;
    }}
</style>
""", unsafe_allow_html=True)


def render_header(with_ticker=False):
    st.markdown(f"""
    <div class="ecng-header">
        <div class="ecng-header-row">
            <div class="ecng-logo-badge"><img src="data:image/png;base64,{LOGO_B64}" alt="ECNG logo" /></div>
            <div>
                <h1>ECNG Energy Group</h1>
                <p>Outstanding terms &middot; pricing &middot; mobile snapshot</p>
            </div>
        </div>
    </div>
    <div class="ecng-header-divider{'' if with_ticker else ' rounded'}"></div>
    """, unsafe_allow_html=True)


def render_ticker(snapshot):
    gas_strip = next((s for s in snapshot.get("pricing", []) if s["name"] == "Gas Strip"), None)
    if not gas_strip or not gas_strip["rows"]:
        return
    # forward term (the one after prompt), since ECNG buys forward — falls
    # back to the first row if there's only one available
    rows = gas_strip["rows"]
    nearest = rows[1] if len(rows) > 1 else rows[0]
    items = "".join(
        f'<span class="item">{h}<b>{nearest["prices"].get(h, "—")}</b></span>'
        for h in gas_strip["hubs"]
    )
    st.markdown(
        f'<div class="ecng-ticker">'
        f'<span class="item" style="color:#8fa5c7;">{nearest["period_label"].upper()}</span>'
        f'{items}</div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------- data loading
@st.cache_data(ttl=300)
def load_snapshot():
    try:
        share_url = st.secrets.get("ONEDRIVE_SHARE_URL", "").strip()
    except Exception:
        share_url = ""
    if share_url:
        try:
            xlsx_bytes = onedrive_source.fetch_excel_bytes(share_url)
            return excel_parser.parse_workbook_bytes(xlsx_bytes)
        except Exception as e:
            st.session_state["_onedrive_error"] = str(e)

    data_path = Path(__file__).parent / "data.json"
    if not data_path.exists():
        return None
    return json.loads(data_path.read_text())


# ---------------------------------------------------------------- helpers
def rep_initials(name):
    if not name:
        return "?"
    parts = str(name).strip().split()
    return (parts[0][:2] if len(parts) == 1 else parts[0][0] + parts[-1][0]).upper()


def fmt_date(v):
    if not v:
        return "—"
    try:
        return datetime.strptime(v, "%Y-%m-%d").strftime("%b %-d, %Y")
    except ValueError:
        return v


def term_str(start, end):
    def short(v):
        try:
            return datetime.strptime(v, "%Y-%m-%d").strftime("%b%y")
        except (ValueError, TypeError):
            return None
    s, e = short(start), short(end)
    if s and e:
        return f"{s}-{e}"
    return s or e or ""


def price_gap(target, market):
    if target is None or market is None:
        return None
    return round(abs(target - market), 2)


def urgency_style(status, target, market):
    """Returns (color, border_width). In the Money is always green and a
    touch thicker, regardless of price gap — it's a status signal, not a
    gap signal. Everything else is colored by how far market has moved
    from target: yellow up to 20c, blue beyond that."""
    if status == "In the Money":
        return GREEN, "6px"
    gap = price_gap(target, market)
    if gap is None:
        return GRAY, "4px"
    if gap > 0.10:
        return NAVY, "4px"   # "blue"
    return GOLD, "4px"       # "yellow" — covers 0 up to 10c


def sort_key(s):
    gap = price_gap(s.get("target"), s.get("market"))
    return gap if gap is not None else float("inf")


def price_str(v):
    return f"${v:,.2f}" if v is not None else "—"


def volume_str(v, product):
    if v is None:
        return "—"
    unit = "GJ" if product and product.strip().lower() == "gas" else "kW"
    return f"{v:,.0f} {unit}"


def delta_str(target, market):
    if target is None or market is None:
        return "—"
    gap = target - market
    if gap > 0:
        return f'<span style="color:{GREEN};font-weight:700;">▼ ${gap:,.2f}</span>'
    if gap < 0:
        return f'<span style="color:{RED};font-weight:700;">▲ ${abs(gap):,.2f}</span>'
    return "$0.00"


def render_deal_card(s):
    is_transacted = s["status"] == "Transacted"
    status_color = STATUS_COLORS.get(s["status"], GRAY)
    is_ab = (s.get("product") or "").strip().lower().startswith("ab")
    volume_row = "" if is_ab else (
        '<div class="row"><div>'
        '<span class="ecng-label">Volume</span>'
        f'<span class="ecng-value">{volume_str(s.get("volume"), s.get("product"))}</span>'
        '</div></div>'
    )

    if is_transacted:
        color = GOLD
        savings = s.get("savings_vs_target")
        if savings is None:
            savings_html = f'<span style="color:{GRAY};">—</span>'
        elif savings < 0:
            savings_html = f'<span style="color:{GREEN};font-weight:700;">▼ Transacted ${abs(savings):,.2f} Below Target</span>'
        elif savings > 0:
            savings_html = f'<span style="color:{RED};font-weight:700;">▲ Transacted ${savings:,.2f} Above Target</span>'
        else:
            savings_html = '<span style="font-weight:700;">Transacted On Target</span>'
        notes = s.get("notes")
        notes_row = (
            f'<div class="row"><div><span class="ecng-label">Notes</span>'
            f'<span class="ecng-value">{notes}</span></div></div>'
        ) if notes else ""
        return (
            f'<div class="ecng-card" style="--accent:{color};">'
            f'<div class="client">{s["client"]}</div>'
            f'<div class="meta"><span class="ecng-avatar">{rep_initials(s["rep"])}</span>'
            f'{s["rep"]} &middot; {s["product"]} &middot; {s.get("delivery_type") or ""}</div>'
            f'{volume_row}'
            '<div class="row" style="align-items:flex-start;"><div>'
            '<span class="ecng-label">Target</span>'
            f'<span class="ecng-value">{price_str(s["target"])}</span>'
            '</div><div style="text-align:right;">'
            '<span class="ecng-label">Transacted Price</span>'
            f'<div class="ecng-value">{price_str(s.get("transacted_price"))}</div>'
            f'<div>{savings_html}</div>'
            '</div></div>'
            '<div class="row" style="align-items:flex-start;"><div>'
            '<span class="ecng-label">Date Transacted</span>'
            f'<span class="ecng-value">{fmt_date(s.get("transacted_date"))}</span>'
            '</div></div>'
            f'{notes_row}'
            '</div>'
        )
    else:
        color, width = urgency_style(s["status"], s.get("target"), s.get("market"))
        days = s.get("days_to_expiry")
        days_str = f"{days}d left" if days is not None else "—"
        return (
            f'<div class="ecng-card" style="--accent:{color}; --accent-width:{width};">'
            f'<div class="client">{s["client"]}</div>'
            f'<div class="meta"><span class="ecng-avatar">{rep_initials(s["rep"])}</span>'
            f'{s["rep"]} &middot; {s["product"]} &middot; {s.get("delivery_type") or ""}</div>'
            f'{volume_row}'
            '<div class="row" style="align-items:flex-start;"><div>'
            '<span class="ecng-label">Target</span>'
            f'<span class="ecng-value">{price_str(s["target"])}</span>'
            '</div><div style="text-align:right;">'
            f'<div style="font-weight:700; color:{DARK_GRAY}; margin-bottom:4px;">{term_str(s["start_date"], s["end_date"])}</div>'
            '<span class="ecng-label">Market</span>'
            f'<div class="ecng-value">{price_str(s["market"])}</div>'
            f'<div>{delta_str(s["target"], s["market"])}</div>'
            '</div></div>'
            '<div class="row" style="align-items:flex-start;"><div>'
            '<span class="ecng-label">Expiry</span>'
            f'<span class="ecng-value">{fmt_date(s["expiry_date"])}</span>'
            '</div>'
            f'<div style="color:{status_color};font-weight:700;">{days_str}</div>'
            '</div></div>'
        )


def render_freshness_banner(snapshot):
    gen_at = snapshot.get("generated_at")
    if not gen_at:
        return
    try:
        gen_dt = datetime.fromisoformat(gen_at)
        age_hours = (datetime.now() - gen_dt).total_seconds() / 3600
        label = gen_dt.strftime("%b %-d, %Y at %-I:%M %p")
        if age_hours > 20:
            st.markdown(
                f'<div class="ecng-stale">⚠️ This snapshot is from {label} — it may be a day or more old.</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(f'<div class="ecng-timestamp-pill">🕐 Snapshot as of {label}</div>', unsafe_allow_html=True)
    except ValueError:
        st.markdown(f'<div class="ecng-timestamp-pill">🕐 Snapshot as of {gen_at}</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------- pages
def go(page):
    st.session_state.page = page


def page_home(snapshot):
    render_header(with_ticker=True)
    render_ticker(snapshot)
    if st.session_state.get("_onedrive_error"):
        st.markdown(
            f'<div class="ecng-stale">⚠️ Could not reach OneDrive — showing the last saved '
            f'snapshot instead.<br><span style="font-size:0.7rem;">{st.session_state["_onedrive_error"]}</span></div>',
            unsafe_allow_html=True,
        )
    render_freshness_banner(snapshot)

    swaps = snapshot.get("swaps", [])
    outstanding_count = sum(1 for s in swaps if s["status"] in ("Active", "In the Money", "Expired"))
    transacted_count = sum(1 for s in swaps if s["status"] == "Transacted")
    pricing_count = len(snapshot.get("pricing", []))

    st.write("")
    with st.container(key="nav_outstanding_wrap"):
        st.button(f"📋  Outstanding  ·  {outstanding_count}", key="nav_outstanding",
                  on_click=go, args=("outstanding",), use_container_width=True)
    st.markdown('<div class="nav-caption">Shows Active + In the Money together, filterable individually</div>', unsafe_allow_html=True)

    with st.container(key="nav_transacted_wrap"):
        st.button(f"💰  Transacted  ·  {transacted_count}", key="nav_transacted",
                  on_click=go, args=("transacted",), use_container_width=True)
    st.markdown('<div class="nav-caption">Deals that have been executed</div>', unsafe_allow_html=True)

    with st.container(key="nav_pricing_wrap"):
        st.button(f"🏷️  Pricing  ·  {pricing_count} curve" + ("s" if pricing_count != 1 else ""),
                  key="nav_pricing", on_click=go, args=("pricing",), use_container_width=True)
    st.markdown('<div class="nav-caption">Current gas and power pricing curves</div>', unsafe_allow_html=True)


def page_outstanding(snapshot):
    with st.container(key="back_home_wrap"):
        st.button("← Home", on_click=go, args=("home",))
    render_header()

    swaps = snapshot.get("swaps", [])
    reps = sorted({s["rep"] for s in swaps if s.get("rep")})
    products = sorted({s["product"] for s in swaps if s.get("product")})

    sel_rep = st.selectbox("My book", ["All reps"] + reps)
    col1, col2 = st.columns(2)
    with col1:
        sel_product = st.selectbox("Product", ["All products"] + products)
    with col2:
        view = st.selectbox("View", ["Active + In the Money", "Active", "In the Money", "Expired"])

    filtered = [s for s in swaps if sel_rep == "All reps" or s["rep"] == sel_rep]
    filtered = [s for s in filtered if sel_product == "All products" or s["product"] == sel_product]
    if view == "Active + In the Money":
        filtered = [s for s in filtered if s["status"] in ("Active", "In the Money")]
    else:
        filtered = [s for s in filtered if s["status"] == view]
    filtered.sort(key=sort_key)

    st.markdown(f"**{len(filtered)}** {view.lower()} deal(s)")

    if not filtered:
        st.markdown('<div class="ecng-empty">📭 &nbsp;Nothing here right now.</div>', unsafe_allow_html=True)
    else:
        for s in filtered:
            st.markdown(render_deal_card(s), unsafe_allow_html=True)


def page_transacted(snapshot):
    with st.container(key="back_home_wrap"):
        st.button("← Home", on_click=go, args=("home",))
    render_header()

    swaps = [s for s in snapshot.get("swaps", []) if s["status"] == "Transacted"]
    reps = sorted({s["rep"] for s in swaps if s.get("rep")})
    products = sorted({s["product"] for s in swaps if s.get("product")})

    sel_rep = st.selectbox("Rep", ["All reps"] + reps)
    sel_product = st.selectbox("Product", ["All products"] + products)

    filtered = [s for s in swaps if sel_rep == "All reps" or s["rep"] == sel_rep]
    filtered = [s for s in filtered if sel_product == "All products" or s["product"] == sel_product]
    filtered.sort(key=lambda s: s.get("transacted_date") or "", reverse=True)

    st.markdown(f"**{len(filtered)}** transacted deal(s)")

    if not filtered:
        st.markdown('<div class="ecng-empty">📭 &nbsp;No transacted deals match the current filters.</div>', unsafe_allow_html=True)
    else:
        for s in filtered:
            st.markdown(render_deal_card(s), unsafe_allow_html=True)


def page_pricing(snapshot):
    with st.container(key="back_home_wrap"):
        st.button("← Home", on_click=go, args=("home",))
    render_header()
    st.subheader("🏷️ Current Pricing")
    st.markdown('<div class="ecng-page-title-divider"></div>', unsafe_allow_html=True)

    for section in snapshot.get("pricing", []):
        st.markdown(f'<div class="ecng-section-pill">{section["name"]}</div>', unsafe_allow_html=True)
        rows = section.get("rows", [])
        if not rows:
            continue
        for r in rows:
            hub_bubbles = "".join(
                f'<div class="ecng-price-card"><div class="hub">{h}</div>'
                f'<div class="price">{r["prices"].get(h, "—").replace("$", "&#36;")}</div></div>'
                for h in section["hubs"]
            )
            st.markdown(
                f'<div class="ecng-curve-row">'
                f'<div class="ecng-price-card ecng-period-badge"><div class="price">{r["period_label"]}</div></div>'
                f'{hub_bubbles}'
                f'</div>',
                unsafe_allow_html=True,
            )


# ---------------------------------------------------------------- main
if "page" not in st.session_state:
    st.session_state.page = "home"

snapshot = load_snapshot()

if not snapshot:
    render_header()
    st.markdown(
        '<div class="ecng-empty">📭 &nbsp;No data yet — waiting on the first snapshot.</div>',
        unsafe_allow_html=True,
    )
    st.stop()

page = st.session_state.page
if page == "home":
    page_home(snapshot)
elif page == "outstanding":
    page_outstanding(snapshot)
elif page == "transacted":
    page_transacted(snapshot)
elif page == "pricing":
    page_pricing(snapshot)
else:
    st.session_state.page = "home"
    st.rerun()
