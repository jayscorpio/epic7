from re import L
from openpyxl import load_workbook
import multiprocessing as mp
from configparser import ConfigParser
import datetime
import statistics

class Hero:
    def __init__(self):
        # Virgin data
        self.items = []
        self.thresholds = {}
    def copy(self):
        new_hero = Hero()
        new_hero.formula    = self.formula.copy()
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
        new_hero.thresholds = self.thresholds.copy()
        return new_hero
    def load(self, dict_data):
        # Load data
        self.formula = dict_data['formula']
        self.thresholds = {}
        if dict_data['CTo'] != 0:
            self.thresholds['CTo'] = dict_data['CTo']
        if dict_data['HTo'] != 0:
            self.thresholds['HTo'] = dict_data['HTo']
        if dict_data['SPo'] != 0:
            self.thresholds['SPo'] = dict_data['SPo']
        if dict_data['TANKo'] != 0:
            self.thresholds['TANKo'] = dict_data['TANKo']
        if dict_data['EVo'] != 0:
            self.thresholds['Evo'] = dict_data['EVo']
        if dict_data['BL'] != 0:
            self.thresholds['BL'] = dict_data['BL']
        if dict_data['FB'] != 0:
            self.thresholds['FB'] = dict_data['FB']
        if dict_data['IM'] != 0:
            self.thresholds['IM'] = dict_data['IM']
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
    def get_formula(self):
        """To Get formula as a list

        Returns:
            list: formula
        """
        return self.formula
    def chk_thresholds(self):
        """To Check thresholds pass or not

        Returns:
            bool: pass thresholds check
        """
        status = self.get_status()
        for criteria, value in self.thresholds.items():
            if criteria == 'CTo':
                if status['CT'] < value:
                    return False
            elif criteria == 'HTo':
                if status['HT'] < value:
                    return False
            elif criteria == 'SPo':
                if status['SP'] < value:
                    return False
            elif criteria == 'TANKo':
                if self.calc_criteria('TANK', status) < value:
                    return False
            elif criteria == 'EVo':
                if status['EV'] < value:
                    return False
            elif criteria == 'BL':
                if not status['BL']:
                    return False
            elif criteria == 'FB':
                if not status['FB']:
                    return False
            elif criteria == 'IM':
                if not status['IM']:
                    return False
        return True
    def strip(self):
        """To remove all clothes
        """
        self.items = []
    def equip(self, item):
        """To equip one item

        Args:
            item (dict): item data
        """
        self.items.append(item)
    def get_status(self):
        """Get hero status

        Returns:
            dict: all kinds of status
        """
        ctr_sets = {
            'AT': 0,
            'DF': 0,
            'HP': 0,
            'SP': 0,
            'CT': 0,
            'CD': 0,
            'HT': 0,
            'EV': 0,
            'FB': 0,
            'IM': 0,
            'BL': 0,
            'FU': 0,
            'TO': 0,
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
        if ctr_sets['CD'] > 3:
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
        # FB
        result['FB'] = (ctr_sets['FB'] > 3)
        # FU
        result['FU'] = (ctr_sets['FU'] > 3)
        # IM
        result['IM'] = (ctr_sets['IM'] > 1)
        # BL
        result['BL'] = (ctr_sets['BL'] > 3)
        return result
    def get_benchmark(self):
        """To calc benchmark

        Returns:
            list: benchmark
        """
        benchmark = []
        status = self.get_status()
        for crtr in self.formula:
            benchmark.append(self.calc_criteria(crtr, status))
        return benchmark
    def calc_criteria(self, criteria, status):
        """To calc on criteria

        Args:
            criteria (str): Criteria
            status (dict): Hero status

        Returns:
            int: mark of the criteria
        """
        if criteria == 'SP':
            return status['SP']
        elif criteria == 'HP':
            return status['HP']
        elif criteria == 'CMGnFU':
            vl = 100 - status['CT'] + status['CT'] * status['CD'] / 100
            return vl
        elif criteria == 'CMG':
            vl = 100 - status['CT'] + status['CT'] * status['CD'] / 100
            if status['FU']:
                return vl * 1.3
            return vl
        elif criteria == 'DMGnFU':
            vl = status['AT'] * (100 - status['CT']) + status['AT'] * status['CT'] * status['CD'] / 100
            return vl
        elif criteria == 'DMG':
            vl = status['AT'] * (100 - status['CT']) + status['AT'] * status['CT'] * status['CD'] / 100
            if status['FU']:
                return vl * 1.3
            return vl
        elif criteria == 'DPSnFU':
            return self.calc_criteria('DMG', status) * status['SP']
        elif criteria == 'DPS':
            vl =  self.calc_criteria('DMG', status) * status['SP']
            if status['FU']:
                return vl * 1.3
            return vl
        elif criteria == 'DPSSnFU':
            return self.calc_criteria('DMG', status) * status['SP'] * status['SP']
        elif criteria == 'DPSS':
            vl = self.calc_criteria('DMG', status) * status['SP'] * status['SP']
            if status['FU']:
                return vl * 1.3
            return vl
        elif criteria == 'HT':
            return status['HT']
        elif criteria == 'CT':
            return status['CT']
        elif criteria == 'EV':
            return status['EV']
        elif criteria == 'TANK':
            return status['HP'] * (1 + status['DF'] / 300)
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
        formula = []
        for i in range(1, 3):
            if rw[i].value != None:
                formula.append(rw[i].value)
        hero.load({
            'formula': formula,
            'CTo':   rw[4].value,
            'HTo':   rw[5].value,
            'SPo':   rw[6].value,
            'TANKo': rw[7].value,
            'EVo':   rw[8].value,
            'BL':    rw[9].value,
            'FB':    rw[10].value,
            'IM':    rw[11].value,            
            'AT':    rw[26].value,
            'ATa':   rw[27].value,
            'ATp':   rw[28].value,
            'DF':    rw[29].value,
            'DFa':   rw[30].value,
            'DFp':   rw[31].value,
            'HP':    rw[32].value,
            'HPa':   rw[33].value,
            'HPp':   rw[34].value,
            'SP':    rw[35].value,
            'SPa':   rw[36].value,
            'CT':    rw[37].value,
            'CTa':   rw[38].value,
            'CD':    rw[39].value,
            'CDa':   rw[40].value,
            'HT':    rw[41].value,
            'HTa':   rw[42].value,
            'EV':    rw[43].value,
            'EVa':   rw[44].value,
        })
        # Check previous item data
        if rw[52].value is not None:
            for itm in items['weapon']:
                if itm['id'] == int(rw[52].value) - 1:
                    hero.equip(itm)
            for itm in items['head']:
                if itm['id'] == int(rw[53].value) - 1:
                    hero.equip(itm)
            for itm in items['armor']:
                if itm['id'] == int(rw[54].value) - 1:
                    hero.equip(itm)
            for itm in items['neck']:
                if itm['id'] == int(rw[55].value) - 1:
                    hero.equip(itm)
            for itm in items['ring']:
                if itm['id'] == int(rw[56].value) - 1:
                    hero.equip(itm)
            for itm in items['shoe']:
                if itm['id'] == int(rw[57].value) - 1:
                    hero.equip(itm)
            # Remove used item from the list
            for itm in items['weapon']:
                if itm['id'] == int(rw[52].value) - 1:
                    items['weapon'].remove(itm)
                    break
            for itm in items['head']:
                if itm['id'] == int(rw[53].value) - 1:
                    items['head'].remove(itm)
                    break
            for itm in items['armor']:
                if itm['id'] == int(rw[54].value) - 1:
                    items['armor'].remove(itm)
                    break
            for itm in items['neck']:
                if itm['id'] == int(rw[55].value) - 1:
                    items['neck'].remove(itm)
                    break
            for itm in items['ring']:
                if itm['id'] == int(rw[56].value) - 1:
                    items['ring'].remove(itm)
                    break
            for itm in items['shoe']:
                if itm['id'] == int(rw[57].value) - 1:
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
def calc_benchmark_group(items, hero, idx, total, queue, priority, flg_debug=False):
    if flg_debug:
        print('calc_benchmark_group called with')
        print('idx={idx}'.format(idx=idx))
        print('total={total}'.format(total=total))
        print('queue={queue}'.format(queue=queue))
        print('priority={priority}'.format(priority=priority))
    result = None
    hero_st = {'AT': 0, 'DF': 0, 'HP': 0, 'SP': 0, 'CT': 0, 'CD': 0, 'HT': 0, 'EV': 0}
    # 1. Use new ones only
    if priority < min([
        len(items['weapon']),
        len(items['head']),
        len(items['armor']),
        len(items['neck']),
        len(items['ring']),
        len(items['shoe']),
    ]):
        items_flattened = {'w': [], 'h': [], 'a': [], 'n': [], 'r': [], 's': []}
        for itm in items['weapon'][priority]:
            items_flattened['w'].append(itm['item'])
        for itm in items['head'][priority]:
            items_flattened['h'].append(itm['item'])
        for itm in items['armor'][priority]:
            items_flattened['a'].append(itm['item'])
        for itm in items['neck'][priority]:
            items_flattened['n'].append(itm['item'])
        for itm in items['ring'][priority]:
            items_flattened['r'].append(itm['item'])
        for itm in items['shoe'][priority]:
            items_flattened['s'].append(itm['item'])
        if flg_debug:
            print('Thread {nm} joins calculating {pos}'.format(
                nm=idx,
                pos=len(items_flattened['w'])*len(items_flattened['h'])*len(items_flattened['a'])*len(items_flattened['n'])*len(items_flattened['r']*len(items_flattened['s']))
            ))
        result = calc_benchmark_group_on_items_set(
            hero, items_flattened, total, idx, result, flg_debug
        )
    # 2. Use new ones to combine old ones
    # New weapons
    if priority < len(items['weapon']):
        items_flattened = {'w': [], 'h': [], 'a': [], 'n': [], 'r': [], 's': []}
        for itm in items['weapon'][priority]:
            items_flattened['w'].append(itm['item'])
        for i in range(priority):
            if i < len(items['head']):
                for itm in items['head'][i]:
                    items_flattened['h'].append(itm['item'])
            if i < len(items['armor']):
                for itm in items['armor'][i]:
                    items_flattened['a'].append(itm['item'])
            if i < len(items['neck']):
                for itm in items['neck'][i]:
                    items_flattened['n'].append(itm['item'])
            if i < len(items['ring']):
                for itm in items['ring'][i]:
                    items_flattened['r'].append(itm['item'])
            if i < len(items['shoe']):
                for itm in items['shoe'][i]:
                    items_flattened['s'].append(itm['item'])
        if flg_debug:
            print('Thread {nm} joins calculating {pos}'.format(
                nm=idx,
                pos=len(items_flattened['w'])*len(items_flattened['h'])*len(items_flattened['a'])*len(items_flattened['n'])*len(items_flattened['r']*len(items_flattened['s']))
            ))
        result = calc_benchmark_group_on_items_set(
            hero, items_flattened, total, idx, result, flg_debug
        )
    # New heads
    if priority < len(items['head']):
        items_flattened = {'w': [], 'h': [], 'a': [], 'n': [], 'r': [], 's': []}
        for itm in items['head'][priority]:
            items_flattened['h'].append(itm['item'])
        for i in range(priority):
            if i < len(items['weapon']):
                for itm in items['weapon'][i]:
                    items_flattened['w'].append(itm['item'])
            if i < len(items['armor']):
                for itm in items['armor'][i]:
                    items_flattened['a'].append(itm['item'])
            if i < len(items['neck']):
                for itm in items['neck'][i]:
                    items_flattened['n'].append(itm['item'])
            if i < len(items['ring']):
                for itm in items['ring'][i]:
                    items_flattened['r'].append(itm['item'])
            if i < len(items['shoe']):
                for itm in items['shoe'][i]:
                    items_flattened['s'].append(itm['item'])
        if flg_debug:
            print('Thread {nm} joins calculating {pos}'.format(
                nm=idx,
                pos=len(items_flattened['w'])*len(items_flattened['h'])*len(items_flattened['a'])*len(items_flattened['n'])*len(items_flattened['r']*len(items_flattened['s']))
            ))
        result = calc_benchmark_group_on_items_set(
            hero, items_flattened, total, idx, result, flg_debug
        )
    # New armors
    if priority < len(items['armor']):
        items_flattened = {'w': [], 'h': [], 'a': [], 'n': [], 'r': [], 's': []}
        for itm in items['armor'][priority]:
            items_flattened['a'].append(itm['item'])
        for i in range(priority):
            if i < len(items['weapon']):
                for itm in items['weapon'][i]:
                    items_flattened['w'].append(itm['item'])
            if i < len(items['head']):
                for itm in items['head'][i]:
                    items_flattened['h'].append(itm['item'])
            if i < len(items['neck']):
                for itm in items['neck'][i]:
                    items_flattened['n'].append(itm['item'])
            if i < len(items['ring']):
                for itm in items['ring'][i]:
                    items_flattened['r'].append(itm['item'])
            if i < len(items['shoe']):
                for itm in items['shoe'][i]:
                    items_flattened['s'].append(itm['item'])
        if flg_debug:
            print('Thread {nm} joins calculating {pos}'.format(
                nm=idx,
                pos=len(items_flattened['w'])*len(items_flattened['h'])*len(items_flattened['a'])*len(items_flattened['n'])*len(items_flattened['r']*len(items_flattened['s']))
            ))
        result = calc_benchmark_group_on_items_set(
            hero, items_flattened, total, idx, result, flg_debug
        )
    # New necks
    if priority < len(items['neck']):
        items_flattened = {'w': [], 'h': [], 'a': [], 'n': [], 'r': [], 's': []}
        for itm in items['neck'][priority]:
            items_flattened['n'].append(itm['item'])
        for i in range(priority):
            if i < len(items['weapon']):
                for itm in items['weapon'][i]:
                    items_flattened['w'].append(itm['item'])
            if i < len(items['head']):
                for itm in items['head'][i]:
                    items_flattened['h'].append(itm['item'])
            if i < len(items['armor']):
                for itm in items['armor'][i]:
                    items_flattened['a'].append(itm['item'])
            if i < len(items['ring']):
                for itm in items['ring'][i]:
                    items_flattened['r'].append(itm['item'])
            if i < len(items['shoe']):
                for itm in items['shoe'][i]:
                    items_flattened['s'].append(itm['item'])
        if flg_debug:
            print('Thread {nm} joins calculating {pos}'.format(
                nm=idx,
                pos=len(items_flattened['w'])*len(items_flattened['h'])*len(items_flattened['a'])*len(items_flattened['n'])*len(items_flattened['r']*len(items_flattened['s']))
            ))
        result = calc_benchmark_group_on_items_set(
            hero, items_flattened, total, idx, result, flg_debug
        )
    # New rings
    if priority < len(items['ring']):
        items_flattened = {'w': [], 'h': [], 'a': [], 'n': [], 'r': [], 's': []}
        for itm in items['ring'][priority]:
            items_flattened['r'].append(itm['item'])
        for i in range(priority):
            if i < len(items['weapon']):
                for itm in items['weapon'][i]:
                    items_flattened['w'].append(itm['item'])
            if i < len(items['head']):
                for itm in items['head'][i]:
                    items_flattened['h'].append(itm['item'])
            if i < len(items['armor']):
                for itm in items['armor'][i]:
                    items_flattened['a'].append(itm['item'])
            if i < len(items['neck']):
                for itm in items['neck'][i]:
                    items_flattened['n'].append(itm['item'])
            if i < len(items['shoe']):
                for itm in items['shoe'][i]:
                    items_flattened['s'].append(itm['item'])
        if flg_debug:
            print('Thread {nm} joins calculating {pos}'.format(
                nm=idx,
                pos=len(items_flattened['w'])*len(items_flattened['h'])*len(items_flattened['a'])*len(items_flattened['n'])*len(items_flattened['r']*len(items_flattened['s']))
            ))
        result = calc_benchmark_group_on_items_set(
            hero, items_flattened, total, idx, result, flg_debug
        )
    # New shoes
    if priority < len(items['shoe']):
        items_flattened = {'w': [], 'h': [], 'a': [], 'n': [], 'r': [], 's': []}
        for itm in items['shoe'][priority]:
            items_flattened['s'].append(itm['item'])
        for i in range(priority):
            if i < len(items['weapon']):
                for itm in items['weapon'][i]:
                    items_flattened['w'].append(itm['item'])
            if i < len(items['head']):
                for itm in items['head'][i]:
                    items_flattened['h'].append(itm['item'])
            if i < len(items['armor']):
                for itm in items['armor'][i]:
                    items_flattened['a'].append(itm['item'])
            if i < len(items['neck']):
                for itm in items['neck'][i]:
                    items_flattened['n'].append(itm['item'])
            if i < len(items['ring']):
                for itm in items['ring'][i]:
                    items_flattened['r'].append(itm['item'])
        if flg_debug:
            print('Thread {nm} joins calculating {pos}'.format(
                nm=idx,
                pos=len(items_flattened['w'])*len(items_flattened['h'])*len(items_flattened['a'])*len(items_flattened['n'])*len(items_flattened['r']*len(items_flattened['s']))
            ))
        result = calc_benchmark_group_on_items_set(
            hero, items_flattened, total, idx, result, flg_debug
        )
    queue.put(result)
def calc_benchmark_group_on_items_set(
    hero, items, thread_num, idx_thread, result_prev=None, flg_debug=False
):
    # By default benchmark parameter number is 3
    if result_prev == None:
        benchmark_best = [0, 0, 0]
        set_best = None
        hero_st = None
    else:
        benchmark_best = result_prev['benchmark_best']
        set_best = result_prev['set_best']
        hero_st = result_prev['hero_st']
    # A counter
    cnt = 0
    # An index for multi processing
    idx_thread_current = thread_num - 1
    for wp in items['w']:
        for hd in items['h']:
            for am in items['a']:
                for nk in items['n']:
                    for rg in items['r']:
                        for sh in items['s']:
                            # Filter for multi processing
                            # Loop index for multi processing
                            if idx_thread_current == thread_num - 1:
                                idx_thread_current = 0
                            else:
                                idx_thread_current += 1
                            if idx_thread == idx_thread_current:
                                cnt += 1
                                hero.strip()
                                hero.equip(wp)
                                hero.equip(hd)
                                hero.equip(am)
                                hero.equip(nk)
                                hero.equip(rg)
                                hero.equip(sh)
                                # Check thresholds pass?
                                if hero.chk_thresholds():
                                    benchmark = hero.get_benchmark()
                                    # Check benchmark valid
                                    if benchmark == None:
                                        continue
                                    # Loop into all parameters
                                    for i in range(len(benchmark)):
                                        # New is better
                                        if benchmark_best[i] < benchmark[i]:
                                            # Use new
                                            benchmark_best = benchmark
                                            if 'id' not in wp:
                                                print('error')
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
    if flg_debug:
        print('actually calculated {num}'.format(num=cnt))
    return {
        'benchmark_best': benchmark_best,
        'set_best': set_best,
        'hero_st': hero_st
    }
def calc_item_score_on_formula(item, formula):
    """Calculate score for the item according to the formula
    The result shall be a list containing powers from each criteria

    Args:
        item ([type]): [description]
        formula ([type]): [description]

    Returns:
        [list]: each element is the score for one criteria
    """
    result = []
    # Loop all criteria
    for criteria in formula:
        # Add the power to result
        result.append(calc_item_score_on_criteria(item, criteria))
    return result
def calc_item_score_on_criteria(item, criteria, result=None):
    """Calculate score for the item according to the criteria
    The result shall be a list containing the min and max power pairs
    Each element in the list shall be a list with min and max power
    The values can be regarded as min powers
    The sets can be regarded as max powers cause it may improve result
    only in some of the cases

    Args:
        item (dict): item data
        criteria (str): criteria used

    Returns:
        list: result containing each criteria score
    """
    if result == None:
        result = []
    if criteria == 'AT':
        # AT
        value = 0
        for attribute in item['attributes']:
            if attribute['type'] == 'ATa':
                value = attribute['value']
        result.append([value, value])
        value = 0
        for attribute in item['attributes']:
           if attribute['type'] == 'ATp':
                value = attribute['value']
        if item['set'] == 'AT':
            result.append([value, value + 35])
        else:
            result.append([value, value])
    elif criteria == 'CT':
        # CT
        value = 0
        for attribute in item['attributes']:
            if attribute['type'] == 'CTa':
                value = attribute['value']
        if item['set'] == 'CT':
            result.append([value, value + 12])
        else:
            result.append([value, value])
    elif criteria == 'CD':
        # CD
        value = 0
        for attribute in item['attributes']:
            if attribute['type'] == 'CDa':
                value = attribute['value']
        if item['set'] == 'CD':
            result.append([value, value + 40])
        else:
            result.append([value, value])
    elif criteria == 'SP':
        # SP
        value = 0
        for attribute in item['attributes']:
            if attribute['type'] == 'SPa':
                value = attribute['value']
        result.append([value, value])
        if item['set'] == 'SP':
            result.append([0, 1])
        else:
            result.append([0, 0])
    elif criteria == 'CMGnFU':
        # CT + CD
        result = calc_item_score_on_criteria(item, 'CT', result)
        result = calc_item_score_on_criteria(item, 'CD', result)
    elif criteria == 'CMG':
        # CT + CD + FU = CMGnFU + FU
        result = calc_item_score_on_criteria(item, 'CMGnFU', result)
        result = calc_item_score_on_criteria(item, 'FU', result)
    elif criteria == 'DMGnFU':
        # CT + CD + AT = CMGnFU + AT
        result = calc_item_score_on_criteria(item, 'CMGnFU', result)
        result = calc_item_score_on_criteria(item, 'AT', result)
    elif criteria == 'DMG':
        # CT + CD + AT + FU = DMGnFU + FU
        result = calc_item_score_on_criteria(item, 'DMGnFU', result)
        result = calc_item_score_on_criteria(item, 'FU', result)
    elif criteria in ('DPSnFU', 'DPSSnFU'):
        # CT + CD + AT + SP = DMGnFU + SP
        result = calc_item_score_on_criteria(item, 'DMGnFU', result)
        result = calc_item_score_on_criteria(item, 'SP', result)
    elif criteria in ('DPS', 'DPSS'):
        # CT + CD + AT + SP + FU = DPSnFU + FU
        result = calc_item_score_on_criteria(item, 'DPSnFU', result)
        result = calc_item_score_on_criteria(item, 'FU', result)
    elif criteria == 'HP':
        # HP
        value = 0
        for attribute in item['attributes']:
            if attribute['type'] == 'HPa':
                value = attribute['value']
        result.append([value, value])
        value = 0
        for attribute in item['attributes']:
            if attribute['type'] == 'HPp':
                value = attribute['value']
        if item['set'] == 'HP':
            result.append([value, value + 15])
        else:
            result.append([value, value])
    elif criteria == 'DF':
        # DF
        value = 0
        for attribute in item['attributes']:
            if attribute['type'] == 'DFa':
                value = attribute['value']
        result.append([value, value])
        value = 0
        for attribute in item['attributes']:
            if attribute['type'] == 'DFp':
                value = attribute['value']
        if item['set'] == 'DF':
            result.append([value, value + 15])
        else:
            result.append([value, value])
    elif criteria == 'TANK':
        # HP, DF
        result = calc_item_score_on_criteria(item, 'HP', result)
        result = calc_item_score_on_criteria(item, 'DF', result)
    elif criteria == 'HT':
        value = 0
        for attribute in item['attributes']:
            if attribute['type'] == 'HTa':
                value = attribute['value']
        # HT
        if item['set'] == 'HT':
            result.append([value, value + 20])
        else:
            result.append([value, value])
    elif criteria == 'EV':
        # EV
        value = 0
        for attribute in item['attributes']:
            if attribute['type'] == 'EVa':
                value = attribute['value']
        if item['set'] == 'EV':
            result.append([value, value + 20])
        else:
            result.append([value, value])
    elif criteria == 'FU':
        # FU
        if item['set'] == 'FU':
            result.append([0, 1])
        else:
            result.append([0, 0])
    else:
        print('unknown criteria {m}'.format(m=criteria))
    return result
def compare_score(score_0, score_1):
    """Compare score
    If at least one pair compare is worse and no is better or not worse or
    unknown, regard as worse
    If at least one pair compare is better and no is worse or not better or
    unknown, regard as better
    Args:
        score_0 (list): compared score
        score_1 (list): comparing score

    Returns:
        str: 'better'/'worse'/'same'/'unknown'
    """
    flg_same = True
    idx = 0
    # Loop into all criteria scores
    for scs in score_0:
        flg_better = False
        flg_worse = False
        flg_not_better = False
        flg_not_worse = False
        flg_unknown = False
        idx_2 = 0
        for sc in scs:
            result = compare_min_max(sc, score_1[idx][idx_2])
            if result != 'same':
                flg_same = False
                if result == 'better':
                    flg_better = True
                elif result == 'worse':
                    flg_worse = True
                elif result == 'not better':
                    flg_not_better = True
                elif result == 'not worse':
                    flg_not_worse = True
                elif result == 'unknown':
                    flg_unknown = True
            idx_2 += 1
        # Better?
        if flg_better and not (flg_worse or flg_not_better or flg_unknown):
            return 'better'
        # Worse?
        if flg_worse and not (flg_better or flg_not_worse or flg_unknown):
            return 'worse'
        # Same?
        if flg_same:
            # Loop to next group
            idx += 1
        # Unknown? Regard as unknown and exit
        break
    # No difference found
    if flg_same:
        return 'same'
    # Otherwise unknown
    return 'unknown'
def compare_min_max(pair_0, pair_1):
    """Compare min max value pairs

    Args:
        pair_0 (list): min max value pairs comparing
        pair_1 (list): min max value pairs compared

    Returns:
        str: better/worse/not better/not worse/unknown
    """
    # Min value better
    if pair_0[0] > pair_1[0]:
        # Max value better
        if pair_0[1] > pair_1[1]:
            return 'better'
        # Max value same
        elif pair_0[1] == pair_1[1]:
            return 'not worse'
        # Max value worse
        else:
            return 'unknown'
    # Min value same
    elif pair_0[0] == pair_1[0]:
        # Max value better
        if pair_0[1] > pair_1[1]:
            return 'not worse'
        # Max value same
        elif pair_0[1] == pair_1[1]:
            return 'same'
        # Max value worse
        else:
            return 'not better'
    # Min value worse
    else:
        # Max value better
        if pair_0[1] > pair_1[1]:
            return 'unknown'
        # Max value same
        elif pair_0[1] == pair_1[1]:
            return 'not better'
        # Max value worse
        else:
            return 'worse'
def priorize_items_by_formula(items, formula, debug=False):
    """Priorize items according to the formula.
    Each criteria is used to filter all the rest candidate items until all
    criteria is checked.
    Only targeted crteria can be counted, not thresholded ones
    The result shall be a list containing groups of items whoes power are equal
    Each group shall be marked with the power of each criteria as a list
    So that the best groups will be processed first to speed up
    If the best groups cannot provide a matched result (due to thresholds),
    Next group shall be used for further iteration

    Args:
        items ([list]): The original items pool
        formula ([list]): consists of criteria

    Returns:
        [dict]: priorized items in 6 classes
    """
    result = []
    # Loop all items in candidates
    for item in items:
        # Calculate this item
        score = calc_item_score_on_formula(item, formula)
        if debug:
            print('-----------------------------------------')
            print('add item with score')
            print(score)
        # Find the matching group
        idx_results = 0
        flg_added = False
        # Loop into all groups from top to bottom
        for rslts in result:
            idx_results += 1
            # Loop all items in a group
            flg_stay = False
            not_beaten_results = []
            beaten_results = []
            idx_result = 0
            for rslt in rslts:
                # Compare
                compare_result = compare_score(score, rslt['score'])
                # New one is same
                if compare_result == 'same':
                    # Directly add into same group
                    result[idx_results - 1].append(
                        {'score': score, 'item': item}
                    )
                    flg_added = True
                    break
                # New one is unknown
                elif compare_result == 'unknown':
                    # Stay in this group
                    flg_stay = True
                    # Add compared as not beaten
                    not_beaten_results.append(rslt)
                # New one is better
                elif compare_result == 'better':
                    # Add compared as beaten
                    beaten_results.append(rslt)
                # New one is worse
                elif compare_result == 'worse':
                    # Loop to next group
                    break
                idx_result += 1
            else:
                # Stay in this group
                if flg_stay:
                    # Refresh this group with not beaten ones
                    result[idx_results - 1] = not_beaten_results
                    # Add into this group
                    result[idx_results - 1].append(
                        {'score': score, 'item': item}
                    )
                    flg_added = True
                    # Beat some of the candidate to a new lower group
                    if len(beaten_results) > 0:
                        result.insert(
                            idx_results, beaten_results
                        )
                # Not stay in this group, means nothing is unknown or same
                else:
                    # This one has beaten someone
                    if len(beaten_results) > 0:
                        # This one has been beaten
                        if len(not_beaten_results) > 0:
                            # This shall not happen
                            print('error')
                        else:
                            # Create a new group ahead
                            result.insert(
                                idx_results - 1, [{'score': score, 'item': item}]
                            )
                            flg_added = True
                    else:
                        # Loop to next group
                        pass                
            if flg_added:
                break
        else:
            result.append([{'score': score, 'item': item}])
        if debug:
            print('the result is now')
            idx = 0
            for rslts in result:
                print('rank {nm}'.format(nm=idx))
                for rslt in rslts:
                    print('{id}{score}'.format(id=rslt['item']['id'], score=rslt['score']))
                idx += 1
    return result

if __name__ == '__main__':
    flg_debug = False
    flg_single_thread = False
    # Open excel data
    cfg = ConfigParser()
    cfg.read('config.ini')
    task_number = int(cfg['CompuPower']['ThreadNumber'])
    pth_data = cfg['Files']['InputData']
    # Speed calculation
    cal_spd = []
    WB = load_workbook(pth_data)
    # Load items data
    items = {}
    # Load each sheet
    load_sheet_item('weapon', items)
    load_sheet_item('head',   items)
    load_sheet_item('armor',  items)
    load_sheet_item('neck',   items)
    load_sheet_item('ring',   items)
    load_sheet_item('shoe',   items)
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
            # Priorize items
            items_priorized = {}
            # Loop all 6 classes to be filtered
            for itm_type, itms in items.items():
                items_priorized[itm_type] = priorize_items_by_formula(
                    itms, formula
                )
            max_loop = max(
                len(items_priorized['weapon']),
                len(items_priorized['head']),
                len(items_priorized['armor']),
                len(items_priorized['neck']),
                len(items_priorized['ring']),
                len(items_priorized['shoe']),
            )
            for i in range(max_loop):
                print('Start calculation round {id}'.format(id=i))
                # Debug with single thread
                if flg_single_thread:
                    hero = hr.copy()
                    q = mp.Queue()
                    calc_benchmark_group(
                        items_priorized, hero, 0, 1, q, i, flg_debug
                    )
                    result_best = q.get()
                    if result_best['set_best'] == None:
                        result_best = None
                # Normal running with multiple threads
                else:
                    processes = []
                    # Queue for results
                    q = mp.Queue()
                    for idx_task in range(task_number):
                        hero = hr.copy()
                        process = mp.Process(
                            target=calc_benchmark_group,
                            args=(
                                items_priorized,
                                hero,
                                idx_task,
                                task_number,
                                q,
                                i,
                                flg_debug
                            )
                        )
                        processes.append(process)
                        process.start()
                    for idx_task in range(task_number):
                        processes[idx_task].join()
                    # Choose best result
                    result_best = None
                    for idx_task in range(task_number):
                        result = q.get()
                        if result['set_best'] != None:
                            for j in range(len(result['benchmark_best'])):
                                # New is better
                                if result_best == None or (
                                    result['benchmark_best'][j] > result_best['benchmark_best'][j]
                                ):
                                    # Update best result
                                    result_best = result
                                    # Stop comparing
                                    break
                                # New is worse
                                elif result['benchmark_best'][j] < result_best['benchmark_best'][j]:
                                    # Give up this result
                                    break
                                # New is same
                                else:
                                    # Compare next parameter
                                    pass
                            else:
                                print('same results detected, please consider add more criteria')
                # Qualified result found
                if result_best != None:
                    # Update excel sheet
                    # Target sheet
                    ws = WB['hero']
                    # Debug
                    # The index used here is incorrect, shall use .id
                    # hero = hr.copy()
                    # hero.strip()
                    # if set_best[0] != None:
                    #     hero.equip(items['weapon'][set_best[0]])
                    # if set_best[1] != None:
                    #     hero.equip(items['head'][set_best[1]])
                    # if set_best[2] != None:
                    #     hero.equip(items['armor'][set_best[2]])
                    # if set_best[3] != None:
                    #     hero.equip(items['neck'][set_best[3]])
                    # if set_best[4] != None:
                    #     print(len(items['ring']))
                    #     print(set_best[4])
                    #     hero.equip(items['ring'][set_best[4]])
                    # if set_best[5] != None:
                    #     hero.equip(items['shoe'][set_best[5]])
                    # benchmark = hero.get_benchmark()
                    # Target row
                    idx_row = idx_hero + 2
                    ws['BA{idx}'.format(idx=idx_row)] = result_best['set_best'][0] + 1
                    ws['BB{idx}'.format(idx=idx_row)] = result_best['set_best'][1] + 1
                    ws['BC{idx}'.format(idx=idx_row)] = result_best['set_best'][2] + 1
                    ws['BD{idx}'.format(idx=idx_row)] = result_best['set_best'][3] + 1
                    ws['BE{idx}'.format(idx=idx_row)] = result_best['set_best'][4] + 1
                    ws['BF{idx}'.format(idx=idx_row)] = result_best['set_best'][5] + 1
                    # Update benchmark for comparing
                    ws['BG{idx}'.format(idx=idx_row)] = result_best['benchmark_best'][0]
                    if len(result_best['benchmark_best']) > 1:
                        ws['BH{idx}'.format(idx=idx_row)] = result_best['benchmark_best'][1]
                        if len(result_best['benchmark_best']) > 2:
                            ws['BI{idx}'.format(idx=idx_row)] = result_best['benchmark_best'][2]
                    ws['BJ{idx}'.format(idx=idx_row)] = result_best['hero_st']['AT']
                    ws['BK{idx}'.format(idx=idx_row)] = result_best['hero_st']['DF']
                    ws['BL{idx}'.format(idx=idx_row)] = result_best['hero_st']['HP']
                    ws['BM{idx}'.format(idx=idx_row)] = result_best['hero_st']['SP']
                    ws['BN{idx}'.format(idx=idx_row)] = result_best['hero_st']['CT']
                    ws['BO{idx}'.format(idx=idx_row)] = result_best['hero_st']['CD']
                    ws['BP{idx}'.format(idx=idx_row)] = result_best['hero_st']['HT']
                    ws['BQ{idx}'.format(idx=idx_row)] = result_best['hero_st']['EV']
                    # Save Excel
                    WB.save(pth_data)
                    # Remove used item from the list
                    for itm in items['weapon']:
                        if itm['id'] == result_best['set_best'][0]:
                            items['weapon'].remove(itm)
                            break
                    for itm in items['head']:
                        if itm['id'] == result_best['set_best'][1]:
                            items['head'].remove(itm)
                            break
                    for itm in items['armor']:
                        if itm['id'] == result_best['set_best'][2]:
                            items['armor'].remove(itm)
                            break
                    for itm in items['neck']:
                        if itm['id'] == result_best['set_best'][3]:
                            items['neck'].remove(itm)
                            break
                    for itm in items['ring']:
                        if itm['id'] == result_best['set_best'][4]:
                            items['ring'].remove(itm)
                            break
                    for itm in items['shoe']:
                        if itm['id'] == result_best['set_best'][5]:
                            items['shoe'].remove(itm)
                            break
                    break
            # tm_delta = datetime.datetime.now() - tm_st
            # print('Used {tm}'.format(tm=tm_delta))
            # if tm_delta.seconds > 10:
            #     spd = ctr_total/tm_delta.seconds
            #     print('Performance: {nm} p/s'.format(
            #         nm=ctr_total/tm_delta.seconds
            #     ))
            #     cal_spd.append(spd)
        idx_hero += 1
