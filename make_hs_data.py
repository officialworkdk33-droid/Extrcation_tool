# Converts UPDATE_HS_CODE.xlsx into hs_data.json (built into the exe by BUILD_EXE.bat)
import danfoss_engine as E
d = E.save_hs_data('UPDATE_HS_CODE.xlsx', 'hs_data.json')
print('HS data built: %d HS codes, %d RPN check codes, columns %s' % (d['count'], len(d['rpn']), d['columns']))
