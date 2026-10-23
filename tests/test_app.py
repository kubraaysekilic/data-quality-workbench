from pathlib import Path
from streamlit.testing.v1 import AppTest

def test_sample_workflow():
    app=AppTest.from_file(str(Path(__file__).parents[1]/'app.py')).run(timeout=30)
    assert not app.exception
    app.toggle[0].set_value(True).run()
    assert not app.exception
    app.multiselect[0].set_value(['email']).run()
    app.multiselect[1].set_value(['date']).run()
    assert app.metric[0].value=='4'
    app.checkbox[0].set_value(True).run()
    app.checkbox[1].set_value(True).run()
    assert not app.exception
