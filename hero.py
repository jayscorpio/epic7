import json
import os

class Database:
    def __init__(self):
        self.data_hero = {}
        for path in os.listdir('gamedatabase/src/hero'):
            self.load_data_hero(os.path.join('gamedatabase/src/hero/', path))
        self.load_data_text('gamedatabase/src/text/text_zhs.json')
        self.load_data_new('new.json')
        self.list_name_id = None
    def load_data_hero(self, path):
        data = json.load(open(path))
        hid = data['id']
        if hid in self.data_hero:
            print('e 10')
            return False
        self.data_hero[hid] = data
    def load_data_text(self, path):
        data = json.load(open(path))
        self.data_text = data
    def load_data_new(self, path):
        data = json.load(open(path))
        self.data_new = data
    def make_hero(self, id, lv, rank, star):
        hero = Hero(id)
        hero.set_lv(lv)
        hero.set_star(star)
        hero.set_data(self.data_hero[id])
        return hero
    def make_hero_new(self, name):
        if name not in self.data_new:
            return False
        hero = Hero(None)
        data = self.data_new[name]
        hero.set_at(data['at'])
        hero.set_df(data['df'])
        hero.set_hp(data['hp'])
        hero.set_sp(data['sp'])
        hero.set_ct(data['ct'])
        hero.set_cd(data['cd'])
        hero.set_ht(data['ht'])
        hero.set_ev(data['ev'])
        return hero
    def make_by_name(self, name, lv, rank, star):
        # List is not ready
        if self.list_name_id == None:
            self.list_name_id = {}
            for key in self.data_text:
                if len(key) != 10:
                    continue
                if key.startswith('chrn_'):
                    self.list_name_id[self.data_text[key]] = key[5:]
        # Find hero id
        if name not in self.list_name_id:
            return self.make_hero_new(name)
        id = self.list_name_id[name]
        if id not in self.data_hero:
            return self.make_hero_new(name)
        return self.make_hero(self.list_name_id[name], lv, rank, star)
        
class Hero:
    def __init__(self, id):
        self.id = id
    def set_star(self, star):
        self.star = star
    def set_lv(self, lv):
        self.lv = lv
    def set_data(self, data):
        self.data = data
    def set_at(self, value):
        self.at = value
    def set_df(self, value):
        self.df = value
    def set_hp(self, value):
        self.hp = value
    def set_sp(self, value):
        self.sp = value
    def set_ct(self, value):
        self.ct = value
    def set_cd(self, value):
        self.cd = value
    def set_ht(self, value):
        self.ht = value
    def set_ev(self, value):
        self.ev = value
    def get_states(self):
        if self.id == None:
            return self.get_states_new()
        result = {}
        result['cp'] = 0
        star_ratio = (1 + (self.star - 1) * 0.075)
        at_b = self.data['stats']['bra'] * 0.6 * (self.lv / 6 + 1) * star_ratio
        at_b = int(at_b)
        df_b = (30 + self.data['stats']['fai'] * 0.3) * (self.lv / 8 + 1) * star_ratio
        df_b = int(df_b)
        hp_b = (50 + self.data['stats']['int'] * 1.4) * (self.lv / 3 + 1) * star_ratio
        hp_b = int(hp_b)
        sp_b = 60 + self.data['stats']['des'] / 1.6
        sp_b = int(sp_b)
        ct_b = 0.15
        cd_b = 1.5
        to_b = 0.05
        ht_b = 0
        ev_b = 0
        # Star
        at_s = 0
        atp_s = 0
        df_s = 0
        dfp_s = 0
        hp_s = 0
        hpp_s = 0
        sp_s = 0
        ct_s = 0
        cd_s = 0
        to_s = 0
        ht_s = 0
        ev_s = 0
        for i in range(self.star):
            for data in self.data['zodiac_tree'][i]['stats']:
                if data['stat'] == 'att':
                    at_s += data['value']
                elif data['stat'] == 'att_rate':
                    atp_s += data['value']
                elif data['stat'] == 'def':
                    df_s += data['value']
                elif data['stat'] == 'def_rate':
                    dfp_s += data['value']
                elif data['stat'] == 'max_hp':
                    hp_s += data['value']
                elif data['stat'] == 'max_hp_rate':
                    hpp_s += data['value']
                elif data['stat'] == 'speed':
                    sp_s += data['value']
                elif data['stat'] == 'cri':
                    ct_s += data['value']
                elif data['stat'] == 'cri_dmg':
                    cd_s += data['value']
                elif data['stat'] == 'acc':
                    ht_s += data['value']
                elif data['stat'] == 'res':
                    ev_s += data['value']
                else:
                    print('unknown data stat')
                    print(data['stat'])
		# 转职加成
        if 'tree' in self.data['specialty_change']:
            for branch in self.data['specialty_change']['tree']:
                for item in branch:
                    for enhancement in item['enhancements']:
                        stat = enhancement['stat']
                        if stat == 'att':
                            at_s += enhancement['value']
                        elif stat == 'att_rate':
                            atp_s += enhancement['value']
                        elif stat == 'def':
                            df_s += enhancement['value']
                        elif stat == 'def_rate':
                            dfp_s += enhancement['value']
                        elif stat == 'max_hp':
                            hp_s += enhancement['value']
                        elif stat == 'max_hp_rate':
                            hpp_s += enhancement['value']
                        elif stat == 'speed':
                            sp_s += enhancement['value']
                        elif stat == 'cri':
                            ct_s += enhancement['value']
                        elif stat == 'cri_dmg':
                            cd_s += enhancement['value']
                        elif stat == 'acc':
                            ht_s += enhancement['value']
                        elif stat == 'res':
                            ev_s += enhancement['value']
                        elif stat == None:
                            pass
                        else:
                            print(stat)
        result['at'] = int(at_b * (1 + atp_s) + at_s)
        result['df'] = int(df_b * (1 + dfp_s) + df_s)
        result['hp'] = int(hp_b * (1 + hpp_s) + hp_s)
        result['sp'] = int(sp_b + sp_s)
        result['ct'] = ct_b + ct_s
        result['cd'] = cd_b + cd_s
        result['to'] = to_b + to_s
        result['ht'] = ht_b + ht_s
        result['ev'] = ev_b + ev_s
        return result
    def get_states_new(self):
        result = {}
        result['at'] = self.at
        result['df'] = self.df
        result['hp'] = self.hp
        result['sp'] = self.sp
        result['ct'] = self.ct
        result['cd'] = self.cd
        result['ht'] = self.ht
        result['ev'] = self.ev
        return result
