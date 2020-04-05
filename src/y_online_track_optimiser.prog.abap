*&---------------------------------------------------------------------*
*& Report Y_ONLINE_TRACK_OPTIMISER
*&---------------------------------------------------------------------*
* In the SAP Online Track every speaker who proposes a session gets
* to pick their first, second and third preffered time slots
* The goal of the optimiser is to fill as many slots as possible, and
* as a secondary goal, try to take account of the speakers preferences
* i.e. give as many as possible their first choice etc..
*&---------------------------------------------------------------------*
REPORT y_online_track_optimiser.
*--------------------------------------------------------------------*
* Global Definitions
*--------------------------------------------------------------------*
TYPES: BEGIN OF g_typ_choices,
         s_name    TYPE string,                            "Speaker Name
         choice_01 TYPE i,
         choice_02 TYPE i,
         choice_03 TYPE i,
         random    TYPE i,
       END OF   g_typ_choices.

TYPES: g_tt_choices TYPE STANDARD TABLE OF g_typ_choices WITH KEY s_name.

TYPES: BEGIN OF g_typ_alv_output,
         slot01 TYPE string,
         slot02 TYPE string,
         slot03 TYPE string,
         slot04 TYPE string,
         slot05 TYPE string,
         slot06 TYPE string,
         slot07 TYPE string,
         slot08 TYPE string,
         slot09 TYPE string,
         slot10 TYPE string,
         slot11 TYPE string,
         slot12 TYPE string,
         slot13 TYPE string,
         slot14 TYPE string,
         slot15 TYPE string,
         slot16 TYPE string,
         slot17 TYPE string,
         slot18 TYPE string,
         slot19 TYPE string,
         slot20 TYPE string,
         slot21 TYPE string,
         slot22 TYPE string,
         slot23 TYPE string,
         slot24 TYPE string,
         filled TYPE i,
         happy  TYPE i,
         total  TYPE i,
         random TYPE i,
       END OF   g_typ_alv_output.

TYPES: g_tt_alv_output TYPE STANDARD TABLE OF g_typ_alv_output WITH EMPTY KEY.

*--------------------------------------------------------------------*
* Class Defintions
*--------------------------------------------------------------------*
CLASS lcl_persistency_layer DEFINITION.

  PUBLIC SECTION.
    METHODS:
      get_data RETURNING VALUE(rt_choices) TYPE g_tt_choices
               RAISING   zcx_excel.

ENDCLASS.                    "lcl_persistency_layer DEFINITION

CLASS lcl_test_class DEFINITION DEFERRED.

CLASS lcl_model DEFINITION FRIENDS lcl_test_class ##CLASS_FINAL.
  PUBLIC SECTION.
    DATA: mt_choices           TYPE g_tt_choices,
          mt_output            TYPE g_tt_alv_output,
          mo_persistency_layer TYPE REF TO lcl_persistency_layer.

    METHODS: constructor IMPORTING io_pers TYPE REF TO lcl_persistency_layer OPTIONAL,
      import_choices,
      prepare_data_for_output,
      strategy_one,
      strategy_two,
      strategy_three,
      strategy_four.

  PRIVATE SECTION.
    METHODS: fill_slot IMPORTING id_choice_no TYPE i
                                 id_name      TYPE string
                       CHANGING  cs_agenda    LIKE LINE OF mt_output
                                 ct_choices   TYPE g_tt_choices,
      calculate_total CHANGING cs_agenda LIKE LINE OF mt_output,
      shuffle_choices CHANGING ct_choices TYPE g_tt_choices.

ENDCLASS.

CLASS lcl_view DEFINITION ##CLASS_FINAL.
  PUBLIC SECTION.
    METHODS: display IMPORTING it_output TYPE g_tt_alv_output.

ENDCLASS.

CLASS lcl_application DEFINITION ##CLASS_FINAL.
  PUBLIC SECTION.
    CLASS-METHODS: main.

ENDCLASS.
*--------------------------------------------------------------------*
* Seelection Screen
*--------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK blk1 WITH FRAME TITLE TEXT-001.

PARAMETERS: p_file TYPE string LOWER CASE MEMORY ID gr8.

PARAMETERS: p_weight TYPE i DEFAULT 75.                    "Weighting towards filling slots

SELECTION-SCREEN END OF BLOCK blk1.

*--------------------------------------------------------------------*
* At Selection-Screen
*--------------------------------------------------------------------*
* Call up Windows open file dialog
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM select_file_path.

*--------------------------------------------------------------------*
* Start-of-Selection
*--------------------------------------------------------------------*
START-OF-SELECTION.
  lcl_application=>main( ).

*--------------------------------------------------------------------*
* Class Implementations
*--------------------------------------------------------------------*
CLASS lcl_persistency_layer IMPLEMENTATION.

  METHOD get_data.
*--------------------------------------------------------------------*
* RETURNING VALUE(rt_choices) TYPE g_tt_choices
* RAISING   zcx_excel.
*--------------------------------------------------------------------*
* Really, we should use ABAP2XLSX to read spreadsheets
* This is a good example of how to extract the contents from an
* EXCEL file and do whatever you want with the data
*--------------------------------------------------------------------*
* Local variables
    DATA: extension TYPE string,
          reader    TYPE REF TO zif_excel_reader,
          column    TYPE zexcel_cell_column VALUE 1,
          row       TYPE int4 VALUE 1,
          choices   LIKE LINE OF rt_choices.

    FIND REGEX '(\.xlsx|\.xlsm)\s*$' IN p_file SUBMATCHES extension.
    TRANSLATE extension TO UPPER CASE.
    CASE extension.

      WHEN '.XLSX'.
        CREATE OBJECT reader TYPE zcl_excel_reader_2007.
        DATA(excel) = reader->load_file( p_file ).

      WHEN '.XLSM'.
        CREATE OBJECT reader TYPE zcl_excel_reader_xlsm.
        excel = reader->load_file( p_file ).

      WHEN OTHERS.
        MESSAGE 'Unsupported filetype'(006) TYPE 'I'.
        RETURN.

    ENDCASE.

    DATA(worksheet)      = excel->get_active_worksheet( ).
    DATA(highest_column) = worksheet->get_highest_column( ).
    DATA(highest_row)    = worksheet->get_highest_row( ).

    WHILE row <= highest_row.
      WHILE column <= highest_column.
        DATA(col_str) = zcl_excel_common=>convert_column2alpha( column ).
        worksheet->get_cell(
          EXPORTING ip_column = col_str
                    ip_row    = row
          IMPORTING ep_value  = DATA(value) ).
        IF column = 1.
          choices-s_name  = value.
        ELSEIF column = 2.
          choices-choice_01 = value.
        ELSEIF column = 3.
          choices-choice_02 = value.
        ELSEIF column = 4.
          choices-choice_03 = value.
        ELSE.
          EXIT."From looking at columns
        ENDIF.
        column = column + 1.
      ENDWHILE.
      APPEND choices TO rt_choices.
      column = 1.
      row    = row + 1.
    ENDWHILE.

  ENDMETHOD.

ENDCLASS.                    "lcl_persistency_layer

CLASS lcl_model IMPLEMENTATION.

  METHOD constructor.

    IF io_pers IS BOUND.
      mo_persistency_layer = io_pers.
    ELSE.
      CREATE OBJECT mo_persistency_layer.
    ENDIF.

  ENDMETHOD.

  METHOD import_choices.

    TRY.
        mt_choices = mo_persistency_layer->get_data( ).
      CATCH zcx_excel.
        MESSAGE 'Oh Dear!'(005) TYPE 'E'.
    ENDTRY.

  ENDMETHOD.

  METHOD prepare_data_for_output.

    strategy_one( ).
    strategy_two( ).
    strategy_three( ).
    strategy_four( ).

    SORT mt_output BY filled DESCENDING
                      happy  DESCENDING.

    DELETE ADJACENT DUPLICATES FROM mt_output COMPARING ALL FIELDS.

  ENDMETHOD.

  METHOD strategy_one.
*--------------------------------------------------------------------*
* Process speakers in the order they applied
*--------------------------------------------------------------------*
* Local Variables
    DATA: agenda LIKE LINE OF mt_output,
          names  TYPE TABLE OF string.

    DATA(choices) = mt_choices[].

    LOOP AT choices ASSIGNING FIELD-SYMBOL(<choice>).
      COLLECT <choice>-s_name INTO names.
    ENDLOOP.

    LOOP AT names ASSIGNING FIELD-SYMBOL(<name>).

      fill_slot( EXPORTING id_choice_no = 1
                           id_name      = <name>
                 CHANGING  cs_agenda    = agenda
                           ct_choices   = choices ).

      fill_slot( EXPORTING id_choice_no = 2
                           id_name      = <name>
                 CHANGING  cs_agenda    = agenda
                           ct_choices   = choices ).

      fill_slot( EXPORTING id_choice_no = 3
                           id_name      = <name>
                 CHANGING  cs_agenda    = agenda
                           ct_choices   = choices ).

    ENDLOOP.

    calculate_total( CHANGING cs_agenda = agenda ).

    INSERT agenda INTO TABLE mt_output.

  ENDMETHOD.

  METHOD strategy_two.
*--------------------------------------------------------------------*
* Here we concentrate on filling the slots, one at a time
*--------------------------------------------------------------------*
* Local Variables
    DATA: agenda       LIKE LINE OF mt_output,
          current_slot TYPE n LENGTH 2,
          slot_field   TYPE c LENGTH 6.

    FIELD-SYMBOLS: <slot_value> TYPE any.

    DATA(choices) = mt_choices[].

    DO 24 TIMES.

      current_slot = current_slot + 1.
      slot_field   = 'SLOT' && current_slot.

      UNASSIGN <slot_value>.
      ASSIGN COMPONENT slot_field OF STRUCTURE agenda TO <slot_value>.
      IF <slot_value> IS NOT ASSIGNED.
        CONTINUE.
      ENDIF.

      LOOP AT choices ASSIGNING FIELD-SYMBOL(<choice>) WHERE choice_01 = current_slot.

        <slot_value>  = <choice>-s_name.                   "Speaker Name
        agenda-filled = agenda-filled + 1.
        agenda-happy  = agenda-happy + 3.

        "The presentation has been allocated, remove record so it is not allocated again
        DELETE choices WHERE s_name = <choice>-s_name.

        EXIT."From Loop

      ENDLOOP.

      IF <slot_value> IS NOT INITIAL.
        CONTINUE."With next slot
      ENDIF.

      LOOP AT choices ASSIGNING <choice> WHERE choice_02 = current_slot.

        <slot_value>  = <choice>-s_name.                   "Speaker Name
        agenda-filled = agenda-filled + 1.
        agenda-happy  = agenda-happy + 2.

        "The presentation has been allocated, renmove record so it is not allocated again
        DELETE choices WHERE s_name = <choice>-s_name.

        EXIT."From Loop

      ENDLOOP.

      IF <slot_value> IS NOT INITIAL.
        CONTINUE."With next slot
      ENDIF.

      LOOP AT choices ASSIGNING <choice> WHERE choice_03 = current_slot.

        <slot_value>  = <choice>-s_name.                   "Speaker Name
        agenda-filled = agenda-filled + 1.
        agenda-happy  = agenda-happy + 1.

        "The presentation has been allocated, remove record so it is not allocated again
        DELETE choices WHERE s_name = <choice>-s_name.

        EXIT."From Loop

      ENDLOOP.

    ENDDO."The 24 available slots

    calculate_total( CHANGING cs_agenda = agenda ).

    INSERT agenda INTO TABLE mt_output.

  ENDMETHOD.

  METHOD strategy_three.
*--------------------------------------------------------------------*
* Re-apply the first two strategies, but sort the speaker table in a
* random order
*--------------------------------------------------------------------*

    DO 20 TIMES.

      shuffle_choices( CHANGING ct_choices = mt_choices ).

      strategy_one( ).
      strategy_two( ).

    ENDDO.

  ENDMETHOD.

  METHOD strategy_four.
*--------------------------------------------------------------------*
* This is loosely based on something called simulated annealing which
* is all about finding the best way to cool metal down
*--------------------------------------------------------------------*
* Local Variables
    DATA: highest_filled TYPE i,
          highest_happy  TYPE i,
          random_line    TYPE i.

    GET TIME.
    DATA(lo_random) = cl_abap_random_int=>create( seed = CONV i( sy-uzeit )
                                                  min  = 1
                                                  max  = 1000 ).

    LOOP AT mt_output ASSIGNING FIELD-SYMBOL(<output>).
      IF <output>-filled GT highest_filled.
        highest_filled = <output>-filled.
      ENDIF.
      IF <output>-happy GT highest_happy.
        highest_happy = <output>-happy.
      ENDIF.
    ENDLOOP.

    DO 30 TIMES.

      DATA(random_number) = lo_random->get_next( ).

      random_line = ( lines( mt_choices ) * random_number ) / 1000.

      READ TABLE mt_choices INTO DATA(random_choice) INDEX random_line.

      IF sy-subrc NE 0.
        CONTINUE."With next attempt
      ENDIF.

      "Move the selected line to the end
      DATA(prior_state) = mt_choices[].
      DELETE mt_choices WHERE s_name = random_choice-s_name.
      INSERT random_choice INTO TABLE mt_choices.

      strategy_one( ).
      strategy_two( ).

*--------------------------------------------------------------------*
* We only keep changes which have improved our best posible result in
* some way. If things did not get better, we revoke the change.
* If you keep doing this forever you would get the optimal result
* but the number of possible combinations means you can only make so
* many attempts. When we have quantum computers the problem will go away...
*--------------------------------------------------------------------*

      "Let us see if that helped
      DATA(did_that_help) = abap_false.
      LOOP AT mt_output ASSIGNING <output> WHERE filled GE highest_filled
                                           AND   happy  GE highest_happy.
        IF <output>-filled GT highest_filled.
          highest_filled = <output>-filled.
          did_that_help  = abap_true.
        ENDIF.
        IF <output>-happy GT highest_happy.
          highest_happy = <output>-happy.
          did_that_help = abap_true.
        ENDIF.
      ENDLOOP.

      IF did_that_help EQ abap_false.
        "Move the line back to where it was
        mt_choices[] = prior_state[].
      ENDIF.

    ENDDO.

  ENDMETHOD.

*--------------------------------------------------------------------*
* Private Section
*--------------------------------------------------------------------*
  METHOD shuffle_choices.
*--------------------------------------------------------------------*
* CHANGING ct_choices TYPE g_tt_choices
*--------------------------------------------------------------------*
    GET TIME.
    DATA(lo_random) = cl_abap_random_int=>create( seed = CONV i( sy-uzeit )
                                                  min  = 1
                                                  max  = 100 ).

    LOOP AT ct_choices ASSIGNING FIELD-SYMBOL(<choice>).
      <choice>-random = lo_random->get_next( ).
    ENDLOOP.

    SORT ct_choices BY random.

  ENDMETHOD.

  METHOD fill_slot.
*--------------------------------------------------------------------*
* IMPORTING id_choice_no TYPE i
*           id_name      TYPE string
* CHANGING  cs_agenda    LIKE LINE OF mt_output
*           ct_choices   TYPE g_tt_choices
*--------------------------------------------------------------------*
* Local Variables
    DATA: choice_number TYPE n LENGTH 2,
          choice_field  TYPE c LENGTH 9,
          slot_number   TYPE n LENGTH 2,
          slot_field    TYPE c LENGTH 6.

    FIELD-SYMBOLS: <choice_value> TYPE any,
                   <slot_value>   TYPE any.

* Preconditions
    CHECK ct_choices[] IS NOT INITIAL.

    LOOP AT ct_choices ASSIGNING FIELD-SYMBOL(<choice>) WHERE s_name = id_name.

      choice_number = id_choice_no.
      choice_field  = 'CHOICE_' && choice_number.

      UNASSIGN <choice_value>.
      ASSIGN COMPONENT choice_field OF STRUCTURE <choice> TO <choice_value>.
      IF <choice_value> IS NOT ASSIGNED.
        CONTINUE.
      ENDIF.

      slot_number = <choice_value>.
      slot_field  = 'SLOT' && slot_number.

      UNASSIGN <slot_value>.
      ASSIGN COMPONENT slot_field OF STRUCTURE cs_agenda TO <slot_value>.
      IF <slot_value> IS NOT ASSIGNED.
        CONTINUE.
      ENDIF.

      IF <slot_value> IS NOT INITIAL.
        "Slot Already Taken
        CONTINUE.
      ENDIF.

      <slot_value>     = <choice>-s_name.                  "Speaker Name
      cs_agenda-filled = cs_agenda-filled + 1.
      cs_agenda-happy  = cs_agenda-happy + ( 4 - id_choice_no ).

      "The presentation has been allocated, renove record so it is not allocated again
      DELETE ct_choices WHERE s_name = <choice>-s_name.

    ENDLOOP.

  ENDMETHOD.

  METHOD calculate_total.
*--------------------------------------------------------------------*
* CHANGING cs_agenda LIKE LINE OF mt_output
*--------------------------------------------------------------------*
* On the selection screen we have a weighting towards filling the slots
* I have defaulted that to 75% as generally it is more important to fill
* all the slots than to keep the presenters happy. However it may be
* possible to do both
*--------------------------------------------------------------------*
    "Maximum filled value is all 24 slots filled
    DATA(filled_value) = ( cs_agenda-filled * 1000 ) / 24. "i.e. maximum = 1000

    "Maximum happy value is 24 x 3 = 72
    DATA(happy_value)  = ( cs_agenda-happy  * 1000 ) / 72. "i.e. maximum = 1000

    DATA(weighted_value) = filled_value * (         p_weight   / 100 ) +
                           happy_value  * ( ( 100 - p_weight ) / 100 ).

    cs_agenda-total = weighted_value / 10.

  ENDMETHOD.

ENDCLASS.

CLASS lcl_view IMPLEMENTATION.

  METHOD display.
*--------------------------------------------------------------------*
*   IMPORTING it_output TYPE g_tt_alv_output.
*--------------------------------------------------------------------*

    DATA(lt_output) = it_output[].                         "Need to convert due to CHANGING

    TRY.
        cl_salv_table=>factory(
          IMPORTING
            r_salv_table = DATA(lo_alv_grid)
          CHANGING
            t_table      = lt_output[] ).

* Application Specific Changes
        DATA(lo_columns) = lo_alv_grid->get_columns( ).
        lo_columns->set_optimize( if_salv_c_bool_sap=>true ).

        DATA(lo_column) = lo_columns->get_column( 'SLOT01' ).
        lo_column->set_long_text( '12 - 13' ).
        lo_column->set_medium_text( '12 - 13' ).
        lo_column->set_short_text( '12 - 13' ).

        lo_column = lo_columns->get_column( 'SLOT02' ).
        lo_column->set_long_text( '13 - 14' ).
        lo_column->set_medium_text( '13 - 14' ).
        lo_column->set_short_text( '13 - 14' ).

        lo_column = lo_columns->get_column( 'SLOT03' ).
        lo_column->set_long_text( '14 - 15' ).
        lo_column->set_medium_text( '14 - 15' ).
        lo_column->set_short_text( '14 - 15' ).

        lo_column = lo_columns->get_column( 'SLOT04' ).
        lo_column->set_long_text( '15 - 16' ).
        lo_column->set_medium_text( '15 - 16' ).
        lo_column->set_short_text( '15 - 16' ).

        lo_column = lo_columns->get_column( 'SLOT05' ).
        lo_column->set_long_text( '16 - 17' ).
        lo_column->set_medium_text( '16 - 17' ).
        lo_column->set_short_text( '16 - 17' ).

        lo_column = lo_columns->get_column( 'SLOT06' ).
        lo_column->set_long_text( '17 - 18' ).
        lo_column->set_medium_text( '17 - 18' ).
        lo_column->set_short_text( '17 - 18' ).

        lo_column = lo_columns->get_column( 'SLOT07' ).
        lo_column->set_long_text( '18 - 19' ).
        lo_column->set_medium_text( '18 - 19' ).
        lo_column->set_short_text( '18 - 19' ).

        lo_column = lo_columns->get_column( 'SLOT08' ).
        lo_column->set_long_text( '19 - 20' ).
        lo_column->set_medium_text( '19 - 20' ).
        lo_column->set_short_text( '19 - 20' ).

        lo_column = lo_columns->get_column( 'SLOT09' ).
        lo_column->set_long_text( '20 - 21' ).
        lo_column->set_medium_text( '20 - 21' ).
        lo_column->set_short_text( '20 - 21' ).

        lo_column = lo_columns->get_column( 'SLOT10' ).
        lo_column->set_long_text( '21 - 22' ).
        lo_column->set_medium_text( '21 - 22' ).
        lo_column->set_short_text( '21 - 22' ).

        lo_column = lo_columns->get_column( 'SLOT11' ).
        lo_column->set_long_text( '22 - 23' ).
        lo_column->set_medium_text( '22 - 23' ).
        lo_column->set_short_text( '22 - 23' ).

        lo_column = lo_columns->get_column( 'SLOT12' ).
        lo_column->set_long_text( '23 - 24' ).
        lo_column->set_medium_text( '23 - 24' ).
        lo_column->set_short_text( '23 - 24' ).

        lo_column = lo_columns->get_column( 'SLOT13' ).
        lo_column->set_long_text( '24 - 01' ).
        lo_column->set_medium_text( '24 - 01' ).
        lo_column->set_short_text( '24 - 01' ).

        lo_column = lo_columns->get_column( 'SLOT14' ).
        lo_column->set_long_text( '01 - 02' ).
        lo_column->set_medium_text( '01 - 02' ).
        lo_column->set_short_text( '01 - 02' ).

        lo_column = lo_columns->get_column( 'SLOT15' ).
        lo_column->set_long_text( '02 - 03' ).
        lo_column->set_medium_text( '02 - 03' ).
        lo_column->set_short_text( '02 - 03' ).

        lo_column = lo_columns->get_column( 'SLOT16' ).
        lo_column->set_long_text( '03 - 04' ).
        lo_column->set_medium_text( '03 - 04' ).
        lo_column->set_short_text( '03 - 04' ).

        lo_column = lo_columns->get_column( 'SLOT17' ).
        lo_column->set_long_text( '04 - 05' ).
        lo_column->set_medium_text( '04 - 05' ).
        lo_column->set_short_text( '04 - 05' ).

        lo_column = lo_columns->get_column( 'SLOT18' ).
        lo_column->set_long_text( '05 - 06' ).
        lo_column->set_medium_text( '05 - 06' ).
        lo_column->set_short_text( '05 - 06' ).

        lo_column = lo_columns->get_column( 'SLOT19' ).
        lo_column->set_long_text( '06 - 07' ).
        lo_column->set_medium_text( '06 - 07' ).
        lo_column->set_short_text( '06 - 07' ).

        lo_column = lo_columns->get_column( 'SLOT20' ).
        lo_column->set_long_text( '07 - 08' ).
        lo_column->set_medium_text( '07 - 08' ).
        lo_column->set_short_text( '07 - 08' ).

        lo_column = lo_columns->get_column( 'SLOT21' ).
        lo_column->set_long_text( '08 - 09' ).
        lo_column->set_medium_text( '08 - 09' ).
        lo_column->set_short_text( '08 - 09' ).

        lo_column = lo_columns->get_column( 'SLOT22' ).
        lo_column->set_long_text( '09 - 10' ).
        lo_column->set_medium_text( '09 - 10' ).
        lo_column->set_short_text( '09 - 10' ).

        lo_column = lo_columns->get_column( 'SLOT23' ).
        lo_column->set_long_text( '10 - 11' ).
        lo_column->set_medium_text( '10 - 11' ).
        lo_column->set_short_text( '10 - 11' ).

        lo_column = lo_columns->get_column( 'SLOT24' ).
        lo_column->set_long_text( '11 - 12' ).
        lo_column->set_medium_text( '11 - 12' ).
        lo_column->set_short_text( '11 - 12' ).

        lo_column = lo_columns->get_column( 'FILLED' ).
        lo_column->set_long_text( 'Filled'(002) ).
        lo_column->set_medium_text( 'Filled'(002) ).
        lo_column->set_short_text( 'Filled'(002) ).

        lo_column = lo_columns->get_column( 'HAPPY' ).
        lo_column->set_long_text( 'Happy'(003) ).
        lo_column->set_medium_text( 'Happy'(003) ).
        lo_column->set_short_text( 'Happy'(003) ).

        lo_column = lo_columns->get_column( 'TOTAL' ).
        lo_column->set_long_text( 'Total'(004) ).
        lo_column->set_medium_text( 'Total'(004) ).
        lo_column->set_short_text( 'Total'(004) ).

        lo_column = lo_columns->get_column( 'RANDOM' ).
        lo_column->set_technical( abap_true ).

        DATA(lo_sorts) = lo_alv_grid->get_sorts( ).

        lo_sorts->add_sort( columnname = 'TOTAL'
                            position   = 1
                            sequence   = if_salv_c_sort=>sort_down         "i.e. descending
                            subtotal   = abap_false ).

* Off we go!
        lo_alv_grid->display( ).

      CATCH cx_salv_msg INTO DATA(lo_salv_msg).
        DATA(ld_error_message) = lo_salv_msg->get_text( ).
        MESSAGE ld_error_message TYPE 'E'.
      CATCH cx_salv_data_error INTO DATA(lo_data_error).
        ld_error_message = lo_data_error->get_text( ).
        MESSAGE ld_error_message TYPE 'E'.
      CATCH cx_salv_existing INTO DATA(lo_existing).
        ld_error_message = lo_existing->get_text( ).
        MESSAGE ld_error_message TYPE 'E'.
      CATCH cx_salv_not_found INTO DATA(lo_not_found).
        "Object = Column
        "Key    = Field Name e.g. VBELN
        ld_error_message = |{ lo_not_found->object } { lo_not_found->key } does not exist|.
        MESSAGE ld_error_message TYPE 'E'.
    ENDTRY.

  ENDMETHOD.

ENDCLASS.

CLASS lcl_application IMPLEMENTATION.

  METHOD main.

    DATA(lo_model) = NEW lcl_model( ).
    DATA(lo_view)  = NEW lcl_view( ).

    lo_model->import_choices( ).
    lo_model->prepare_data_for_output( ).
    lo_view->display( lo_model->mt_output ).

  ENDMETHOD.

ENDCLASS.
*--------------------------------------------------------------------*
* Test Classes & Test Doubles
*--------------------------------------------------------------------*
* Test Doubles
*--------------------------------------------------------------------*
CLASS ltd_persistency_layer DEFINITION INHERITING FROM lcl_persistency_layer ##CLASS_FINAL.

  PUBLIC SECTION.
    METHODS: get_data REDEFINITION.

ENDCLASS.

CLASS ltd_persistency_layer IMPLEMENTATION.

  METHOD get_data.

    rt_choices = VALUE #(
    ( s_name = 'HARDY' choice_01 = 13 choice_02 = 2  choice_03 = 12 )
    ( s_name = 'GOAT'  choice_01 = 4  choice_02 = 5  choice_03 = 6  )
    ( s_name = 'FISH'  choice_01 = 4  choice_02 = 5  choice_03 = 6 )
    ( s_name = 'SHEEP' choice_01 = 20 choice_02 = 22 choice_03 = 24 ) ).

  ENDMETHOD.

ENDCLASS.
*--------------------------------------------------------------------*
* Test Classes
*--------------------------------------------------------------------*
CLASS lcl_test_class DEFINITION FOR TESTING
   RISK LEVEL HARMLESS
   DURATION SHORT
   FINAL.

  PUBLIC SECTION.

  PRIVATE SECTION.
    DATA: mo_class_under_test TYPE REF TO lcl_model.

    METHODS: setup,
*--------------------------------------------------------------------*
* Specifications
*--------------------------------------------------------------------*
      "IT SHOULD.....................
      do_strategy_one        FOR TESTING,
      do_strategy_two        FOR TESTING,
      calculate_total_points FOR TESTING,
      calculate_max_points   FOR TESTING.
    "User Acceptance Tests
    "Helper Methods

ENDCLASS.                    "lcl_test_class DEFINITION

CLASS lcl_test_class IMPLEMENTATION.

  METHOD setup.

    DATA(lo_mock_pers_layer) = NEW ltd_persistency_layer( ).

    CREATE OBJECT mo_class_under_test
      EXPORTING
        io_pers = lo_mock_pers_layer.

    mo_class_under_test->import_choices( ).

  ENDMETHOD.

  METHOD do_strategy_one.

* Given
* Input Data is created during SETUP

* When
    mo_class_under_test->strategy_one( ).

* Then
    DATA(agenda) = mo_class_under_test->mt_output[ 1 ].

    IF sy-subrc NE 0 OR
       agenda   IS INITIAL.
      DATA(no_error) = abap_false.
    ELSE.
      no_error = abap_true.
    ENDIF.

    cl_abap_unit_assert=>assert_equals( msg = 'Agenda is not filled out'
                                        exp = abap_true
                                        act = no_error ).

    cl_abap_unit_assert=>assert_equals( msg = 'Agenda filled count is blank'
                                        exp = abap_true
                                        act = xsdbool( agenda-filled > 0 ) ).

  ENDMETHOD.

  METHOD do_strategy_two.

* Given
* Input Data is created during SETUP

* When
    mo_class_under_test->strategy_two( ).

* Then
    DATA(agenda) = mo_class_under_test->mt_output[ 1 ].

    IF sy-subrc NE 0 OR
       agenda   IS INITIAL.
      DATA(no_error) = abap_false.
    ELSE.
      no_error = abap_true.
    ENDIF.

    cl_abap_unit_assert=>assert_equals( msg = 'Agenda is not filled out'
                                        exp = abap_true
                                        act = no_error ).

    cl_abap_unit_assert=>assert_equals( msg = 'Agenda filled count is blank'
                                        exp = abap_true
                                        act = xsdbool( agenda-filled > 0 ) ).

  ENDMETHOD.

  METHOD calculate_total_points.

* Given
    p_weight     = 75.
    DATA(agenda) = VALUE g_typ_alv_output( filled = 5
                                           happy  = 15 ).

* When
    mo_class_under_test->calculate_total( CHANGING cs_agenda = agenda ).

* Then
    cl_abap_unit_assert=>assert_equals( msg = 'Points are not calculated correctly'
                                        exp = 21
                                        act = agenda-total ).

  ENDMETHOD.

  METHOD calculate_max_points.

* Given
    p_weight     = 75.
    DATA(agenda) = VALUE g_typ_alv_output( filled = 24
                                           happy  = 72 ).

* When
    mo_class_under_test->calculate_total( CHANGING cs_agenda = agenda ).

* Then
    cl_abap_unit_assert=>assert_equals( msg = 'Points are not calculated correctly'
                                        exp = 100
                                        act = agenda-total ).

  ENDMETHOD.

ENDCLASS.
*&---------------------------------------------------------------------*
*& Form SELECT_FILE_PATH
*&---------------------------------------------------------------------*
*& User Dialog to Choose a File from their local PC
*& This comes from ZDEMO_EXCEL37
*&---------------------------------------------------------------------*
FORM select_file_path.
* Local Variables
  DATA: fields      TYPE dynpread_tabtype,
        field       LIKE LINE OF fields,
        files       TYPE filetable,
        file_filter TYPE string.

  DATA(repid) = sy-repid.

  CALL FUNCTION 'DYNP_VALUES_READ'
    EXPORTING
      dyname               = repid
      dynumb               = '1000'
      request              = 'A'
    TABLES
      dynpfields           = fields
    EXCEPTIONS
      invalid_abapworkarea = 01
      invalid_dynprofield  = 02
      invalid_dynproname   = 03
      invalid_dynpronummer = 04
      invalid_request      = 05
      no_fielddescription  = 06
      undefind_error       = 07.

  IF sy-subrc NE 0.
    RETURN.
  ENDIF.

  READ TABLE fields INTO field WITH KEY fieldname = 'P_PFILE'.
  p_file = field-fieldvalue.

  file_filter = 'Excel Files (*.XLSX;*.XLSM)|*.XLSX;*.XLSM' ##NO_TEXT.
  cl_gui_frontend_services=>file_open_dialog( EXPORTING
                                                default_filename = p_file
                                                file_filter      = file_filter
                                              CHANGING
                                                file_table       = files
                                                rc               = sy-tabix
                                              EXCEPTIONS
                                                OTHERS           = 1 ).
  IF sy-subrc NE 0.
    RETURN.
  ENDIF.

  READ TABLE files INDEX 1 INTO p_file.

ENDFORM.
