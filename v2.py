from openpyxl import load_workbook
import multiprocessing as mp
from configparser import ConfigParser
import datetime

class Hero:
    def __init__(self):
        # Virgin data
        self.items = []
    def copy(self):
        new_hero = Hero()
        new_hero.formula    = self.formula
        new_hero.AT         = self.AT
        new_hero.ATp        = self.ATp
        new_hero.ATa        = self.ATa
        new_hero.DF         = self.DF
        new_hero.DFp        = self.DFp
        new_hero.DFa        = self.DFa
        new_hero.HP         = self.HP
        new_hero.HPp        = self.HPp
        new_hero.HPa        = self.HPa
        new_hero.SP         = self.SP
        new_hero.SPa        = self.SPa
        new_hero.CT         = self.CT
        new_hero.CTa        = self.CTa
        new_hero.CD         = self.CD
        new_hero.CDa        = self.CDa
        new_hero.HT         = self.HT
        new_hero.HTa        = self.HTa
        new_hero.EV         = self.EV
        new_hero.EVa        = self.EVa
        new_hero.TO         = self.TO
        new_hero.TOa        = self.TOa
        new_hero.FB         = self.FB
        new_hero.FBa        = self.FBa
        return new_hero
    def load(self, dict_data):
        # Load data
        self.formula = dict_data['formula'].split(';')
        self.AT = dict_data['AT']
        self.ATp = dict_data['ATp']
        self.ATa = dict_data['ATa']
        self.DF = dict_data['DF']
        self.DFp = dict_data['DFp']
        self.DFa = dict_data['DFa']
        self.HP = dict_data['HP']
        self.HPp = dict_data['HPp']
        self.HPa = dict_data['HPa']
        self.SP = dict_data['SP']
        self.SPa = dict_data['SPa']
        self.CT = dict_data['CT']
        self.CTa = dict_data['CTa']
        self.CD = dict_data['CD']
        self.CDa = dict_data['CDa']
        self.HT = dict_data['HT']
        self.HTa = dict_data['HTa']
        self.EV = dict_data['EV']
        self.EVa = dict_data['EVa']
        self.TO = dict_data['TO']
        self.TOa = dict_data['TOa']
        self.FB = dict_data['FB']
        self.FBa = dict_data['FBa']
    '''
    To Get formula in list
    '''
    def get_formula(self):
        return self.formula
    # To remove all clothes
    def strip(self):
        self.items = []
    # To equip one item
    def equip(self, item):
        self.items.append(item)
    def get_status(self):
        ctr_sets = {
            'AT': 0,
            'DF': 0,
            'HP': 0,
            'SP': 0,
            'CT': 0,
            'CD': 0,
            'HT': 0,
            'EV': 0,
            'TO': 0,
            'FB': 0,
            'IM': 0,
            'BL': 0,
            'FU': 0,
        }
        data = {
            'AT': self.AT,
            'ATp': self.ATp,
            'ATa': self.ATa,
            'DF': self.DF,
            'DFp': self.DFp,
            'DFa': self.DFa,
            'HP': self.HP,
            'HPp': self.HPp,
            'HPa': self.HPa,
            'SP': self.SP,
            'SPa': self.SPa,
            'CT': self.CT,
            'CTa': self.CTa,
            'CD': self.CD,
            'CDa': self.CDa,
            'HT': self.HT,
            'HTa': self.HTa,
            'EV': self.EV,
            'EVa': self.EVa,
            'TO': self.TO,
            'TOa': self.TOa,
            'FB': self.FB,
            'FBa': self.FBa,
        }
        for itm in self.items:
            ctr_sets[itm['set']] += 1
            for attrb in itm['attributes']:
                data[attrb['type']] += attrb['value']
        result = {}
        # AT
        vl = data['AT']
        vl_p = data['ATp']
        vl_a = data['ATa']
        if ctr_sets['AT'] > 3:
            vl_p += 35
        result['AT'] = vl * (100 + vl_p) / 100 + vl_a
        # DF
        vl = data['DF']
        vl_p = data['DFp']
        vl_a = data['DFa']
        if ctr_sets['DF'] > 1:
            vl_p += 15
            if ctr_sets['DF'] > 3:
                vl_p += 15
                if ctr_sets['DF'] > 5:
                    vl_p += 15
        result['DF'] = vl * (100 + vl_p) / 100 + vl_a
        # HP
        vl = data['HP']
        vl_p = data['HPp']
        vl_a = data['HPa']
        if ctr_sets['HP'] > 1:
            vl_p += 15
            if ctr_sets['HP'] > 3:
                vl_p += 15
                if ctr_sets['HP'] > 5:
                    vl_p += 15
        result['HP'] = vl * (100 + vl_p) / 100 + vl_a
        # SP
        vl = data['SP']
        vl_a = data['SPa']
        if ctr_sets['SP'] > 3:
            vl_p = 25
        else:
            vl_p = 0
        result['SP'] = vl * (100 + vl_p) / 100 + vl_a
        # CT
        vl = data['CT'] + data['CTa']
        if ctr_sets['CT'] > 1:
            vl += 12
            if ctr_sets['CT'] > 3:
                vl += 12
                if ctr_sets['CT'] > 5:
                    vl += 12
        if vl > 100:
            vl = 100
        result['CT'] = vl
        # CD
        vl = data['CD'] + data['CDa']
        if ctr_sets['CT'] > 3:
            vl += 40
        result['CD'] = vl
        # HT
        vl = data['HT'] + data['HTa']
        if ctr_sets['HT'] > 1:
            vl += 20
            if ctr_sets['HT'] > 3:
                vl += 20
                if ctr_sets['HT'] > 5:
                    vl += 20
        result['HT'] = vl
        # EV
        vl = data['EV'] + data['EVa']
        if ctr_sets['EV'] > 1:
            vl += 20
            if ctr_sets['EV'] > 3:
                vl += 20
                if ctr_sets['EV'] > 5:
                    vl += 20
        result['EV'] = vl
        # TO
        vl = data['TO'] + data['TOa']
        if ctr_sets['TO'] > 1:
            vl += 4
            if ctr_sets['TO'] > 3:
                vl += 4
                if ctr_sets['TO'] > 5:
                    vl += 4
        result['TO'] = vl
        # FB
        result['FB'] = (ctr_sets['FB'] > 3)
        # FU
        result['FU'] = (ctr_sets['FU'] > 3)
        # IM
        result['IM'] = (ctr_sets['IM'] > 1)
        # BL
        result['BL'] = (ctr_sets['BL'] > 3)
        return result
    # To calc benchmark
    def get_benchmark(self):
        benchmark = []
        status = self.get_status()
        for crtr in self.formula:
            benchmark.append(self.calc_criteria(crtr, status))
        return benchmark
    # To calc on criteria
    def calc_criteria(self, criteria, status):
        if criteria == 'SP':
            return status['SP']
        elif criteria == 'HP':
            return status['HP']
        elif criteria == 'DMG':
            vl = status['AT'] * (100 - status['CT']) + status['AT'] * status['CT'] * status['CD'] / 100
            if status['FU']:
                return vl * 1.3
            return vl
        elif criteria == 'DPS':
            return self.calc_criteria('DMG', status) * status['SP']
        elif criteria == 'HT':
            return status['HT']
        elif criteria == 'CT':
            return status['CT']
        elif criteria == 'FB':
            if status['FB']:
                return 1
            return 0
        elif criteria == 'TANK':
            return status['HP'] * (1 + status['DF'] / 300)
        elif criteria == 'BL':
            if status['BL']:
                return 1
            return 0
        elif criteria == 'CMG':
            vl = 100 - status['CT'] + status['CT'] * status['CD'] / 100
            if status['FU']:
                return vl * 1.3
            return vl
        elif criteria == 'DF':
            vl = status['DF']
            return vl
        else:
            print('unexpected criteria {nm}'.format(nm=criteria))
def load_sheet_item(_type, _holder):
    _holder[_type] = []
    ws = WB[_type]
    idx = 0
    for rw in ws.rows:
        itm = {}
        # Index
        itm['id'] = idx
        itm['set'] = rw[0].value
        itm['attributes'] = []
        for i in range(5):
            if rw[i * 2 + 1].value == None:
                continue
            itm['attributes'].append(
                {'type': rw[i * 2 + 1].value, 'value': rw[i * 2 + 2].value}
            )
        _holder[_type].append(itm)
        # Increase index
        idx += 1
def load_sheet_hero(holder, items):
    ws = WB['hero']
    idx = 0
    for rw in ws.rows:
        if idx == 0:
            idx += 1
            continue
        # Check valid line
        if rw[1].value is None:
            continue
        hero = Hero()
        hero.load({
            'formula': rw[1].value,
            'AT': rw[2].value,
            'ATp': rw[4].value,
            'ATa': rw[3].value,
            'DF': rw[5].value,
            'DFp': rw[7].value,
            'DFa': rw[6].value,
            'HP': rw[8].value,
            'HPp': rw[10].value,
            'HPa': rw[9].value,
            'SP': rw[11].value,
            'SPa': rw[12].value,
            'CT': rw[13].value,
            'CTa': rw[14].value,
            'CD': rw[15].value,
            'CDa': rw[16].value,
            'HT': rw[17].value,
            'HTa': rw[18].value,
            'EV': rw[19].value,
            'EVa': rw[20].value,
            'TO': rw[21].value,
            'TOa': rw[22].value,
            'FB': 0,
            'FBa': 0,
        })
        # Check previous item data
        if rw[23].value is not None:
            for itm in items['weapon']:
                if itm['id'] == int(rw[23].value) - 1:
                    hero.equip(itm)
            for itm in items['head']:
                if itm['id'] == int(rw[24].value) - 1:
                    hero.equip(itm)
            for itm in items['armor']:
                if itm['id'] == int(rw[25].value) - 1:
                    hero.equip(itm)
            for itm in items['neck']:
                if itm['id'] == int(rw[26].value) - 1:
                    hero.equip(itm)
            for itm in items['ring']:
                if itm['id'] == int(rw[27].value) - 1:
                    hero.equip(itm)
            for itm in items['shoe']:
                if itm['id'] == int(rw[28].value) - 1:
                    hero.equip(itm)
            # Remove used item from the list
            for itm in items['weapon']:
                if itm['id'] == int(rw[23].value) - 1:
                    items['weapon'].remove(itm)
                    break
            for itm in items['head']:
                if itm['id'] == int(rw[24].value) - 1:
                    items['head'].remove(itm)
                    break
            for itm in items['armor']:
                if itm['id'] == int(rw[25].value) - 1:
                    items['armor'].remove(itm)
                    break
            for itm in items['neck']:
                if itm['id'] == int(rw[26].value) - 1:
                    items['neck'].remove(itm)
                    break
            for itm in items['ring']:
                if itm['id'] == int(rw[27].value) - 1:
                    items['ring'].remove(itm)
                    break
            for itm in items['shoe']:
                if itm['id'] == int(rw[28].value) - 1:
                    items['shoe'].remove(itm)
                    break
        holder.append(hero)
        idx += 1
def wear(_piece, _sum):
    # Add set
    _sum['set'][_piece['set']] += 1
    # Add attributes
    for attr in _piece['attributes']:
        _sum[attr['type']] += int(attr['value'])
def calc_benchmark_group(items, hero, idx, total, queue):
    # By default benchmark parameter number is 3
    benchmark_best = [0, 0, 0]
    set_best = None
    hero_st = {'AT': 0, 'DF': 0, 'HP': 0, 'SP': 0, 'CT': 0, 'CD': 0, 'HT': 0, 'EV': 0}
    idx_wp = 0
    for wp in items['weapon']:
        # Fileter by index
        if idx_wp % total != idx:
            idx_wp += 1
            continue
        for hd in items['head']:
            for am in items['armor']:
                for nk in items['neck']:
                    for rg in items['ring']:
                        for sh in items['shoe']:
                            hero.strip()
                            hero.equip(wp)
                            hero.equip(hd)
                            hero.equip(am)
                            hero.equip(nk)
                            hero.equip(rg)
                            hero.equip(sh)
                            benchmark = hero.get_benchmark()
                            # Check benchmark valid
                            if benchmark == None:
                                continue
                            # print('[{a},{b},{c},{d},{e},{f},{g}]'.format(a=idx_wp,b=idx_hd,c=idx_am,d=idx_nk,e=idx_rg,f=idx_sh,g=benchmark))
                            # Loop into all parameters
                            for i in range(len(benchmark)):
                                # New is better
                                if benchmark_best[i] < benchmark[i]:
                                    # Use new
                                    benchmark_best = benchmark
                                    if 'id' not in wp:
                                        print(wp)
                                    set_best = [
                                        wp['id'],
                                        hd['id'],
                                        am['id'],
                                        nk['id'],
                                        rg['id'],
                                        sh['id'],
                                    ]
                                    hero_st = {
                                        'AT': hero.get_status()['AT'],
                                        'DF': hero.get_status()['DF'],
                                        'HP': hero.get_status()['HP'],
                                        'SP': hero.get_status()['SP'],
                                        'CT': hero.get_status()['CT'],
                                        'CD': hero.get_status()['CD'],
                                        'HT': hero.get_status()['HT'],
                                        'EV': hero.get_status()['EV'],
                                    }
                                # New is worse
                                elif benchmark_best[i] > benchmark[i]:
                                    # Next please
                                    break
                                # New is same
                                else:
                                    # Go to next parameter
                                    pass
        idx_wp += 1
    queue.put({
        'set_best': set_best,
        'benchmark_best': benchmark_best,
        'hero_st': hero_st
    })
def calc_item_score_on_formula(item, formula):
    result = {}
    for criteria in formula:
        if criteria == 'CT':
            if item['set'] == 'CT':
                result['CT set'] = 1
            else:
                result['CT set'] = 0
            for attribute in item['attributes']:
                if attribute['type'] == 'CTa':
                    result['CTa'] = attribute['value']
        elif criteria == 'DPS':
            if item['set'] == 'CT':
                result['CT set'] = 1
            else:
                result['CT set'] = 0
            if item['set'] == 'AT':
                result['AT set'] = 1
            else:
                result['AT set'] = 0
            if item['set'] == 'CD':
                result['CD set'] = 1
            else:
                result['CD set'] = 0
            if item['set'] == 'FU':
                result['FU set'] = 1
            else:
                result['FU set'] = 0
            if item['set'] == 'SP':
                result['SP set'] = 1
            else:
                result['SP set'] = 0
            for attribute in item['attributes']:
                if attribute['type'] == 'CTa':
                    result['CTa'] = attribute['value']
                elif attribute['type'] == 'ATa':
                    result['ATa'] = attribute['value']
                elif attribute['type'] == 'CDa':
                    result['CDa'] = attribute['value']
                elif attribute['type'] == 'ATp':
                    result['ATp'] = attribute['value']
                elif attribute['type'] == 'SPa':
                    result['SPa'] = attribute['value']
        elif criteria == 'SP':
            if item['set'] == 'SP':
                result['SP set'] = 1
            else:
                result['SP set'] = 0
            for attribute in item['attributes']:
                if attribute['type'] == 'SPa':
                    result['SPa'] = attribute['value']
        elif criteria == 'HP':
            if item['set'] == 'HP':
                result['HP set'] = 1
            else:
                result['HP set'] = 0
            for attribute in item['attributes']:
                if attribute['type'] == 'HPa':
                    result['HPa'] = attribute['value']
                if attribute['type'] == 'HPp':
                    result['HPp'] = attribute['value']
        elif criteria == 'DMG':
            if item['set'] == 'CT':
                result['CT set'] = 1
            else:
                result['CT set'] = 0
            if item['set'] == 'AT':
                result['AT set'] = 1
            else:
                result['AT set'] = 0
            if item['set'] == 'CD':
                result['CD set'] = 1
            else:
                result['CD set'] = 0
            if item['set'] == 'FU':
                result['FU set'] = 1
            else:
                result['FU set'] = 0
            for attribute in item['attributes']:
                if attribute['type'] == 'CTa':
                    result['CTa'] = attribute['value']
                elif attribute['type'] == 'ATa':
                    result['ATa'] = attribute['value']
                elif attribute['type'] == 'CDa':
                    result['CDa'] = attribute['value']
                elif attribute['type'] == 'ATp':
                    result['ATp'] = attribute['value']
        elif criteria == 'HT':
            if item['set'] == 'HT':
                result['HT set'] = 1
            else:
                result['HT set'] = 0
            for attribute in item['attributes']:
                if attribute['type'] == 'HTa':
                    result['HTa'] = attribute['value']
        elif criteria == 'FB':
            if item['set'] == 'FB':
                result['FB set'] = 1
            else:
                result['FB set'] = 0
        elif criteria == 'TANK':
            if item['set'] == 'HP':
                result['HP set'] = 1
            else:
                result['HP set'] = 0
            if item['set'] == 'DF':
                result['DF set'] = 1
            else:
                result['DF set'] = 0
            for attribute in item['attributes']:
                if attribute['type'] == 'HPa':
                    result['HPa'] = attribute['value']
                if attribute['type'] == 'HPp':
                    result['HPp'] = attribute['value']
                if attribute['type'] == 'DFa':
                    result['DFa'] = attribute['value']
                if attribute['type'] == 'DFp':
                    result['DFp'] = attribute['value']
        elif criteria == 'BL':
            if item['set'] == 'BL':
                result['BL set'] = 1
            else:
                result['BL set'] = 0
        elif criteria == 'CMG':
            if item['set'] == 'CT':
                result['CT set'] = 1
            else:
                result['CT set'] = 0
            if item['set'] == 'CD':
                result['CD set'] = 1
            else:
                result['CD set'] = 0
            if item['set'] == 'FU':
                result['FU set'] = 1
            else:
                result['FU set'] = 0
            for attribute in item['attributes']:
                if attribute['type'] == 'CTa':
                    result['CTa'] = attribute['value']
                elif attribute['type'] == 'CDa':
                    result['CDa'] = attribute['value']
        elif criteria == 'DF':
            if item['set'] == 'DF':
                result['DF set'] = 1
            else:
                result['DF set'] = 0
            for attribute in item['attributes']:
                if attribute['type'] == 'DFa':
                    result['DFa'] = attribute['value']
                elif attribute['type'] == 'DFp':
                    result['DFp'] = attribute['value']
        else:
            print('unknown criteria {m}'.format(m=criteria))
    return result
'''
To filter items that contributes nothing to the formula.
'''
def filter_items_by_formula(items, formula):
    result = {}
    for itype, dt in items.items():
        # Prepare marks
        marks = {}
        # Prepare good indexes
        idxs_good = set()
        for item in dt:
            # Calculate this item
            mark = calc_item_score_on_formula(item, formula)
            # Compare with previous items
            flg_beaten = False
            idxs_beaten = set()
            # Loop all previous items
            for idx_mark_prev, mark_prev in marks.items():
                flg_any_better = False
                # This one is better than previous one ?
                for ky, vl in mark.items():
                    if ky in mark_prev:
                        base = mark_prev[ky]
                    else:
                        base = 0
                    if vl > base:
                        flg_any_better = True
                        break
                # Beat this one?
                if not flg_any_better:
                    # Beat it
                    flg_beaten = True
                    # Stop comparing
                    break
                flg_any_worse = False
                # Previous one is better than this one?
                for ky, vl in mark_prev.items():
                    if ky in mark:
                        base = mark[ky]
                    else:
                        base = 0
                    if vl > base:
                        flg_any_worse = True
                        break
                # Beat previous one?
                if not flg_any_worse:
                    # Beat the previous one
                    idxs_beaten.add(idx_mark_prev)
            # Remove beaten ones
            for i in idxs_beaten:
                marks.pop(i)
                # Remove it from good indexes
                idxs_good.remove(i)
            # This one is not beaten
            if not flg_beaten:
                # Add into marks
                marks[item['id']] = mark
                # Add into good indexes
                idxs_good.add(item['id'])
            # This one is beaten
            else:
                pass
        # Prepare result type
        result[itype] = []
        for itm in dt:
            if itm['id'] in idxs_good:
                result[itype].append(itm)
    return result
if __name__ == '__main__':
    # Open excel data
    cfg = ConfigParser()
    cfg.read('config.ini')
    task_number = int(cfg['CompuPower']['ThreadNumber'])
    pth_data = cfg['Files']['InputData']
    WB = load_workbook(pth_data)
    # Load items data
    items = {}
    # Load each sheet
    load_sheet_item('weapon', items)
    load_sheet_item('head', items)
    load_sheet_item('armor', items)
    load_sheet_item('neck', items)
    load_sheet_item('ring', items)
    load_sheet_item('shoe', items)
    # Load hero data
    heroes = []
    load_sheet_hero(heroes, items)
    idx_hero = 0
    for hr in heroes:
        # No previous data, re-calculate
        if len(hr.items) == 0:
            # Get formula
            formula = hr.get_formula()
            # Check formula valid
            if len(formula) == 0:
                continue
            items_filtered = filter_items_by_formula(items, formula)
            ctr_total = len(items_filtered['weapon'])*len(items_filtered['head'])*len(items_filtered['armor'])*len(items_filtered['neck'])*len(items_filtered['ring'])*len(items_filtered['shoe'])
            print('Calculate {nm} possibilities'.format(nm=ctr_total))
            tm_st =datetime.datetime.now()
            # Debug with single thread
            if False:
                hero = hr.copy()
                q = mp.Queue()
                calc_benchmark_group(items_filtered, hero, 0, 1, q)
            # Normal running with multiple threads
            else:
                processes = []
                # Queue for results
                q = mp.Queue()
                for i in range(task_number):
                    hero = hr.copy()
                    process = mp.Process(
                        target=calc_benchmark_group,
                        args=(items_filtered, hero, i, task_number, q)
                    )
                    processes.append(process)
                    process.start()
                for i in range(task_number):
                    processes[i].join()
                # Choose best result
                benchmark_best = [0, 0, 0]
                set_best = None
                for i in range(task_number):
                    result = q.get()
                    if result['set_best'] != None:
                        for j in range(len(result['benchmark_best'])):
                            # New is better
                            if result['benchmark_best'][j] > benchmark_best[j]:
                                benchmark_best = result['benchmark_best']
                                set_best = result['set_best']
                                hero_st = result['hero_st']
                            # New is worse
                            elif result['benchmark_best'][j] < benchmark_best[j]:
                                # Give up this result
                                break
                            # New is same
                            else:
                                # Compare next parameter
                                pass
                # Update excel sheet
                if set_best != None:
                    # Target sheet
                    ws = WB['hero']
                    # Target row
                    idx_row = idx_hero + 2
                    ws['X{idx}'.format(idx=idx_row)] = set_best[0] + 1
                    ws['Y{idx}'.format(idx=idx_row)] = set_best[1] + 1
                    ws['Z{idx}'.format(idx=idx_row)] = set_best[2] + 1
                    ws['AA{idx}'.format(idx=idx_row)] = set_best[3] + 1
                    ws['AB{idx}'.format(idx=idx_row)] = set_best[4] + 1
                    ws['AC{idx}'.format(idx=idx_row)] = set_best[5] + 1
                    # Update benchmark for comparing
                    if len(benchmark_best) > 0:
                        ws['AD{idx}'.format(idx=idx_row)] = benchmark_best[0]
                        if len(benchmark_best) > 1:
                            ws['AE{idx}'.format(idx=idx_row)] = benchmark_best[1]
                            if len(benchmark_best) > 2:
                                ws['AF{idx}'.format(idx=idx_row)] = benchmark_best[2]
                    ws['AG{idx}'.format(idx=idx_row)] = hero_st['AT']
                    ws['AH{idx}'.format(idx=idx_row)] = hero_st['DF']
                    ws['AI{idx}'.format(idx=idx_row)] = hero_st['HP']
                    ws['AJ{idx}'.format(idx=idx_row)] = hero_st['SP']
                    ws['AK{idx}'.format(idx=idx_row)] = hero_st['CT']
                    ws['AL{idx}'.format(idx=idx_row)] = hero_st['CD']
                    ws['AM{idx}'.format(idx=idx_row)] = hero_st['HT']
                    ws['AN{idx}'.format(idx=idx_row)] = hero_st['EV']
                    # Save Excel
                    WB.save(pth_data)
                    # Remove used item from the list
                    for itm in items['weapon']:
                        if itm['id'] == set_best[0]:
                            items['weapon'].remove(itm)
                            break
                    for itm in items['head']:
                        if itm['id'] == set_best[1]:
                            items['head'].remove(itm)
                            break
                    for itm in items['armor']:
                        if itm['id'] == set_best[2]:
                            items['armor'].remove(itm)
                            break
                    for itm in items['neck']:
                        if itm['id'] == set_best[3]:
                            items['neck'].remove(itm)
                            break
                    for itm in items['ring']:
                        if itm['id'] == set_best[4]:
                            items['ring'].remove(itm)
                            break
                    for itm in items['shoe']:
                        if itm['id'] == set_best[5]:
                            items['shoe'].remove(itm)
                            break
                    # Equip item on the hero
                    # hr.strip()
                    # hr.equip(items['weapon'][set_best[0]])
                    # hr.equip(items['head'][set_best[1]])
                    # hr.equip(items['armor'][set_best[2]])
                    # hr.equip(items['neck'][set_best[3]])
                    # hr.equip(items['ring'][set_best[4]])
                    # hr.equip(items['shoe'][set_best[5]])
                    # Print result
                    print(benchmark_best)
                    # print(items['weapon'][set_best[0]])
                    # print(items['head'][set_best[1]])
                    # print(items['armor'][set_best[2]])
                    # print(items['neck'][set_best[3]])
                    # print(items['ring'][set_best[4]])
                    # print(items['shoe'][set_best[5]])
            tm_delta = datetime.datetime.now() - tm_st
            print('Used {tm}'.format(tm=tm_delta))
            print('Performance: {nm} p/s'.format(nm=ctr_total/tm_delta.seconds))
        idx_hero += 1
