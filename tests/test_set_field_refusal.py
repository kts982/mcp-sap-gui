"""set_field / set_batch_fields say why a write was refused."""

from .field_fakes import FakeField, make_controller

FIELD = "wnd[0]/usr/txtTHOLT-KTEXT"
OTHER = "wnd[0]/usr/txtTHOLT-LTEXT"
PASSWORD = "wnd[0]/usr/pwdRSYST-BCODE"


def test_set_field_too_long_is_refused_before_writing():
    field = FakeField(text="old", max_length=10)
    c = make_controller({FIELD: field})

    result = c.set_field(FIELD, "Ognissanti 1")  # 12 characters

    assert result == {
        "field_id": FIELD,
        "error": "value is 12 characters, field takes 10",
        "max_length": 10,
        "value_length": 12,
        "status": "rejected",
    }
    assert field.writes == 0
    assert field.text == "old"


def test_set_field_not_changeable_is_refused_before_writing():
    field = FakeField(text="old", changeable=False)
    c = make_controller({FIELD: field})

    result = c.set_field(FIELD, "x")

    assert result == {"field_id": FIELD, "error": "field is not changeable",
                      "status": "rejected"}
    assert field.writes == 0


def test_set_field_that_fits_is_unchanged():
    field = FakeField(max_length=10)
    c = make_controller({FIELD: field})

    result = c.set_field(FIELD, "x" * 10)

    assert result == {"field_id": FIELD, "value": "x" * 10, "status": "success"}
    assert field.text == "x" * 10


def test_set_field_unreadable_properties_do_not_refuse():
    # Neither property readable: write as before, and a failure keeps the
    # old generic message.
    field = FakeField(max_length=3, unreadable=("MaxLength", "Changeable"))
    c = make_controller({FIELD: field})

    assert c.set_field(FIELD, "abc")["status"] == "success"
    assert c.set_field(FIELD, "abcd") == {"field_id": FIELD,
                                          "error": "Could not set field"}


def test_set_field_zero_max_length_is_not_enforced():
    field = FakeField(max_length=0)
    c = make_controller({FIELD: field})
    # MaxLength 0 means "not known", never "takes nothing".
    assert c._write_refusal(field, FIELD, "anything") is None


def test_set_field_password_length_is_not_echoed():
    field = FakeField(max_length=8)
    c = make_controller({PASSWORD: field})

    result = c.set_field(PASSWORD, "x" * 12)

    assert result["status"] == "rejected"
    assert "12" not in result["error"]
    assert "value_length" not in result
    assert result["max_length"] == 8


def test_batch_explains_a_too_long_field():
    long_field = FakeField(max_length=10)
    ok_field = FakeField(max_length=30)
    c = make_controller({FIELD: long_field, OTHER: ok_field})

    result = c.set_batch_fields({FIELD: "Ognissanti 1", OTHER: "All Saints"})

    assert result["succeeded"] == 1
    assert result["failed"] == 1
    assert result["results"] == {
        FIELD: "error: value is 12 characters, field takes 10",
    }


def test_batch_lookup_failure_does_not_reuse_previous_element():
    # The first field is found and set; the second is not found. Without the
    # per-field reset, the explanation would read the first field's element.
    first = FakeField(max_length=2)
    c = make_controller({FIELD: first})
    missing = "wnd[0]/usr/txtNOT-THERE"

    result = c.set_batch_fields({FIELD: "ab", missing: "abc"})

    assert result["results"] == {missing: "error: Could not set field"}


def test_batch_success_path_reads_no_extra_properties():
    class CountingField(FakeField):
        reads = 0

        @property
        def MaxLength(self):  # noqa: N802 (SAP GUI COM name)
            CountingField.reads += 1
            return 10

    field = CountingField()
    c = make_controller({FIELD: field})

    assert c.set_batch_fields({FIELD: "short"})["failed"] == 0
    assert CountingField.reads == 0

