import hashlib
from pathlib import Path
import streamlit as st
from quality import load_table,sheets,inspect,transform,csv_bytes,xlsx_bytes

st.set_page_config(page_title='Data Quality Workbench',page_icon='◈',layout='wide')
st.markdown("""<style>.stApp{background:#f6f8fc}h1{letter-spacing:-1.8px;color:#183350}.block-container{max-width:1300px;padding-top:2.5rem}div[data-testid="stMetric"]{background:white;border:1px solid #e1e8f1;padding:20px;border-radius:12px}section[data-testid="stSidebar"]{background:#edf2f9}div[data-testid="stFileUploader"]{background:#fff;border-radius:12px;padding:16px}</style>""",unsafe_allow_html=True)
with st.sidebar:
    st.markdown('## ◈ Quality Lab')
    st.caption('VERİ KALİTESİ ÇALIŞMA ALANI')
    st.divider()
    st.markdown('**01** Dosya ve şema\n\n**02** Sorunları incele\n\n**03** Düzelt ve indir')
    st.divider()
    st.info('Dosyalar uygulamanın çalıştığı makinede işlenir; diske kaydedilmez. Yerel kullanım önerilir.')
    use_sample=st.toggle('Örnek veriyle dene',value=False)
st.caption('PANDAS + STREAMLIT / DATA WORKBENCH')
st.title('Temiz veri. Güvenilir başlangıç.')
st.write('Eksikleri ve biçim sorunlarını gör. Yalnızca seçtiğin düzeltmeleri uygula.')
uploaded=st.file_uploader('CSV veya Excel dosyanı seç',type=['csv','xlsx'],help='En fazla 10 MB, 50.000 satır, 150 sütun. CSV: UTF-8. İlk satır başlık olmalı.')
if uploaded:
    content=uploaded.getvalue();name=uploaded.name
elif use_sample:
    content=(Path(__file__).parent/'examples/sample.csv').read_bytes();name='sample.csv'
else:
    st.info('Başlamak için dosya yükle veya sol menüden örnek veriyi aç.')
    st.stop()
signature=hashlib.sha256(content).hexdigest()
try:
    sheet=st.selectbox('Excel sayfası',sheets(content)) if name.lower().endswith('.xlsx') else 0
    delimiter=st.selectbox('CSV ayırıcı',[',',';','\t'],format_func=lambda x:{',':'Virgül',';':'Noktalı virgül','\t':'Sekme'}[x]) if name.lower().endswith('.csv') else ','
    frame=load_table(content,name,sheet,delimiter)
except Exception as error:
    st.error(f'Dosya okunamadı: {error}');st.stop()
st.caption(f'{name} · {len(frame):,} satır · {len(frame.columns)} sütun')
if name.lower().endswith('.xlsx'): st.info('Excel içe/dışa aktarımı seçilen sayfanın değerlerini yeni tabloya taşır. Formüller, biçimler, grafikler ve diğer sayfalar korunmaz. Formül önbelleği olmayan hücreler boş görünebilir.')
with st.expander('Doğrulama şeması — e-posta ve tarih sütunlarını seç',expanded=True):
    c1,c2,c3=st.columns(3)
    emails=c1.multiselect('E-posta sütunları',frame.columns,key='emails-'+signature+str(sheet))
    dates=c2.multiselect('Tarih sütunları',frame.columns,key='dates-'+signature+str(sheet))
    date_format=c3.selectbox('Beklenen tarih biçimi',['%Y-%m-%d','%d.%m.%Y','%d/%m/%Y'],format_func=lambda x:{'%Y-%m-%d':'2026-10-23 (YYYY-MM-DD)','%d.%m.%Y':'23.10.2026 (GG.AA.YYYY)','%d/%m/%Y':'23/10/2026 (GG/AA/YYYY)'}[x])
issues=inspect(frame,emails,dates,date_format)
m1,m2,m3,m4=st.columns(4)
m1.metric('Toplam bulgu',len(issues));m2.metric('Eksik hücre',int((issues['Sorun']=='Eksik hücre').sum()))
m3.metric('Yinelenen satır',int(frame.duplicated().sum()));m4.metric('Biçim sorunu',int(issues['Sorun'].isin(['Tarih biçimi','E-posta biçimi']).sum()))
tab1,tab2,tab3=st.tabs(['Sorun raporu','Orijinal veri','Düzeltme ve dışa aktarım'])
with tab1:
    st.caption('Satır numaraları başlık dahil dosya satırını gösterir. E-posta kontrolü yalnızca biçimi doğrular.')
    if issues.empty:st.success('Seçilen kurallara göre sorun bulunmadı.')
    else:st.dataframe(issues,use_container_width=True,hide_index=True)
    st.download_button('Sorun raporunu indir',csv_bytes(issues),'quality-report.csv','text/csv')
with tab2:st.dataframe(frame,use_container_width=True,hide_index=True)
with tab3:
    st.subheader('Düzeltmeleri sen seç')
    a,b=st.columns(2)
    trim=a.checkbox('Hücre başı/sonu boşluklarını temizle')
    drop=b.checkbox('Yinelenen satırları kaldır (ilkini tut)')
    normalize=a.checkbox('Geçerli tarihleri YYYY-MM-DD biçimine dönüştür')
    columns=b.multiselect('Eksikleri doldurulacak sütunlar',frame.columns)
    fill=st.text_input('Eksik hücreler için değer',value='Belirtilmedi')
    result=transform(frame,trim=trim,drop_duplicates=drop,fill_columns=columns,fill_value=fill,date_columns=dates,date_format=date_format,normalize_dates=normalize)
    st.caption('Önizleme seçtiğin işlemlere göre güncellenir. Orijinal dosya değiştirilmez. Geçersiz tarihler otomatik tahmin edilmez.')
    st.dataframe(result,use_container_width=True,hide_index=True)
    st.write(f'Çıktı: **{len(result)} satır** · kaldırılan: **{len(frame)-len(result)}**')
    st.caption('CSV çıktısında formül olarak yorumlanabilecek metinlerin başına güvenli kesme işareti eklenir. XLSX hücreleri metin olarak yazılır.')
    x,y=st.columns(2)
    x.download_button('Düzeltilmiş CSV indir',csv_bytes(result),'cleaned-data.csv','text/csv',type='primary')
    y.download_button('Düzeltilmiş Excel indir',xlsx_bytes(result),'cleaned-data.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
