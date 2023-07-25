from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.clock import Clock
from kivy.core.window import Window
import requests
import json


class MainApp(App):
    def build(self):
        layout = BoxLayout(orientation='vertical')
        self.user_input_skuId = TextInput(hint_text='Enter the skuId value', size_hint_y=0.1)
        submit_button = Button(text='Submit', on_release=self.send_request, size_hint_y=0.1)
        layout.add_widget(self.user_input_skuId)
        layout.add_widget(submit_button)

        result_layout = BoxLayout(orientation='vertical', size_hint_y=None)
        result_layout.bind(minimum_height=result_layout.setter('height')) 

        self.result = Label(font_name='gulim.ttc', size_hint_y=None)
        spacer = Label(size_hint_y=None, height=160)

        result_layout.add_widget(spacer) 
        result_layout.add_widget(self.result) 

        self.scrollview = ScrollView(size_hint_y=0.8, do_scroll_x=False, do_scroll_y=True)
        self.scrollview.add_widget(result_layout)  
        layout.add_widget(self.scrollview)

        return layout

    def send_request(self, instance):
        self.scrollview.do_scroll_y = False

        url = 'https://api.louisvuitton.com/eco-eu/search-merch-eapi/v1/kor-kr/stores/query'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
            'Content-Type': 'application/json',
            'Accept': 'application/json, text/plain, */*',
            'client_secret': '60bbcdcD722D411B88cBb72C8246a22F',
            'client_id': '607e3016889f431fb8020693311016c9',
            'Origin': 'https://kr.louisvuitton.com',
            'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
            'Cookie': 'ATG_SESSION_ID=E4i7k6N4uqdHA3we5jdEXtRl.front131-prd; ATGID=anonymous; anonymous_session=true; _dynSessConf=8923089313706214252; JSESSIONID=E4i7k6N4uqdHA3we5jdEXtRl.front131-prd; SGID=sb.springboot131-prd; bm_sz=C3631D7AA73F951FA5D4AD28574EDBAE~YAAQzTVDF/efuJ+IAQAAvrCnrhQzSqHopKHCkEzH4CDBpVSjt2CdcrK7xKD4Ur9r2WMDl2abOsq90ekwmhfXa7QEKxaH1Mk+tSyR4CGeG4q4OCSHFDM9pFimOHLrKfn5EA4a3tCLGxLgLpHfAfL739DE+UQnjZhg69ZeUaZzpRDauJvEKVUVGz6EGAUKGSmW5dSBtI4CCE+p42IFeiRRd/BI0wsYOdwZksOa0NYERiZ6T5u/Rru2R1t6RtpgvIqHB7ByYZsT7fFJJoNVgmWEvWJI24G6lg8qooPgIyXB2P9hGq/FvRsd170=~3552308~3556914; AKA_A2=A; OPTOUTMULTI=0:0%7Cc1:0%7Cc2:0%7Cc4:0%7Cc3:0; lv-dispatch=kor-kr; lv-dispatch-url=https://kr.louisvuitton.com/kor-kr/products/puzzle-flower-monogram-keyring-s00-nvprod4170081v/M01207; ak_cc=KR; ak_bmsc=2F6DAD582D758BD2C3313B28ECC46CCB~000000000000000000000000000000~YAAQzTVDF9OguJ+IAQAA5rWnrhQ3hb2TxLVo5Ex8aajU2QlZT0aikdEl2ssz2cFI0nmQ/GxtYrC3mQocuAQcOCenI5fdJGx8nlN+TdxWnwr/wYMqkV88UK8+gkHKRkw3HdCs7CutvEHjB8w/PdRr+UnXY0qqzsg+nlFxn7n3arUqMiiuP34Cx8+enjzQYWnz0mbqMJFJXJsy+Qc2EmrNuLHj75DEvsNsWwM9Zluteq80yFHayEf6g6C7F/+z2iQZSy92xgotTC2+OA31Iloncux93WZbgRuUHwVWVPyUV0yg2xhBXogJdYY6Ou2TX/l6eftOFnyJzM4Yo0L3adcFohgE9dnTsRY52jHsnSdsYkhVoMJJ6K4UH/uxIUq9lubHZ+iQ3KrNohQN8dk+VIZYQPi40wOo0cu179k3fAB8zxEi7XgzxA5p/66UW6BmA1XtQR3e5dslAddoTp4vZbQ5vRG4hr8SSrtmuCseU97ONf2f; RT="z=1&dm=louisvuitton.com&si=66447985-e7aa-49f7-a5dc-104845852594&ss=liskqdmw&sl=0&tt=0&bcn=%2F%2F684d0d4b.akstat.io%2F"; _cs_mk=0.926590451296496_1686557407779; _fbp=fb.1.1686557408151.960999036; _scid=196b6e5d-818b-4da9-89b4-b5afe7433432; _scid_r=196b6e5d-818b-4da9-89b4-b5afe7433432; _tt_enable_cookie=1; _ttp=i9BoCdSWrv1ulrGKB-ATML_cHrj; _qubitTracker=gkuisyi8co8-0liskqi6w-69tgrt8; qb_generic=:Yiup8Lm:.louisvuitton.com; _gid=GA1.2.1423735359.1686557411; _ga=GA1.1.795833456.1686557411; _gcl_au=1.1.722436059.1686557411; _abck=49A08A3DFA90539A85752575479FAB37~0~YAAQxjVDF4u9OZ+IAQAA/MynrgrzawyY+shhkV6LdQuVG0VXX4iT6zbtcyMsC/Id3nLEQdA1ueCy6lDYrgFj5UoYbINNklZWNY2c1wCjDBwpnz0nCSN08lvY3+oVS7Jiaaq/+0veidDHwvlkz1ODuOyxXZniHRZGh1xiGdp2TzIE8UVXenBS/FzA3ex++VI7KM5DvizNHFO2QXmt9wDmG6yqPc4Rzo4c2EH8WGTx5zurnhQ2q6vEFWPsTvELpzT0lSZqJUz5qW+H+H2fDZkLgpDZwnzqwfTzLjEp+ICBBIt6LahxoINAZL5WTkE+ks2Q+2cgOUMLcsMYO2bUN8t16QCOSPm7yeBcetP4umorik8tHlA+A4ZGdg/BoOG3QUeNJZxMB0Bf49ZH6u4yKrk5lsY6xi9cvOZYo6xI7+s9~-1~-1~-1; qb_permanent=gkuisyi8co8-0liskqi6w-69tgrt8:1:1:1:1:0::0:1:0:BkhtLj:BkhtLj:::::211.52.1.133:seoul:2261:south%20korea:KR:37.56:127:jung-gu%20seoul:410014:seoul-teukbyeolsi:25025:migrated|1686557413582:EbFa==B=CRhL=s&Fv2m==B=CrkC=MQ&F4TI==B=Cs1j=J5::Yiup9VG:Yiup8WB:0:0:0::0:0:.louisvuitton.com:0; qb_session=1:1:32:EbFa=B&Fv2m=B&F4TI=B:0:Yiup8WB:0:0:0:0:.louisvuitton.com; _sctr=1%7C1686495600000; _ga_S6ED35NYJQ=GS1.1.1686557410.1.0.1686557415.55.0.0; utag_main=v_id:0188aea7b40c00231535fc13a6180506f006406700bd0$_sn:1$_se:8$_ss:0$_st:1686559215620$ses_id:1686557406222%3Bexp-session$_pn:1%3Bexp-session$dc_visit:1$dc_event:5%3Bexp-session$dc_region:eu-central-1%3Bexp-session'
        }
        data = {
            "flagShip": False,
            "country": "KR",
            "latitudeCenter": "37.48654195241737",
            "longitudeCenter": "127.0971466",
            "latitudeA": "37.48953011783969",
            "longitudeA": "127.09046253513489",
            "latitudeB": "37.48355378699504",
            "longitudeB": "127.10383066486511",
            "query": "",
            "clickAndCollect": False,
            "pageType": "productsheet"
        }
        data["skuId"] = self.user_input_skuId.text

        response = requests.post(url, headers=headers, data=json.dumps(data))

        self.result.text = ''
        self.scrollview.scroll_y = 1

        if response.status_code == 200:
            response_data = json.loads(response.content.decode('utf-8'))
            stores = response_data.get('hits', [])
            result_text = ""
            if len(stores) == 0:
                result_text = "No stores found with stock availability for the given skuId."
            else:
                for store in stores:
                    additional_properties = store.get('additionalProperty', [])
                    stock_available = any(
                        item for item in additional_properties if item.get('name') == 'stockAvailability' and item.get('value') == 'true')
                    if stock_available:
                        result_text += f"\nStore Name: {store.get('name')}\n"
            if result_text == "":
                result_text = "No stores found with stock availability for the given skuId."
            self.result.text = result_text
        else:
            self.result.text = f"Error occurred: {response.status_code}\nResponse content: {response.content.decode('utf-8')}\n"
        
        self.user_input_skuId.text = ''

        Clock.schedule_once(self.enable_scroll, 0.1) 

    def enable_scroll(self, dt):
        self.scrollview.do_scroll_y = True  


if __name__ == '__main__':
    MainApp().run()
