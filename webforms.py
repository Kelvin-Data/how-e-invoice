from flask_wtf import FlaskForm
from wtforms import (StringField, PasswordField, SubmitField, BooleanField,
                     EmailField, IntegerField, TextAreaField, TelField,
                     SearchField, DateTimeField, DecimalField, fields,
                     SelectField, DateField)
from wtforms.validators import (DataRequired, EqualTo, Email, InputRequired,
                 Length, Optional)
from wtforms.widgets import Input
from datetime import datetime

class EndUserLicenseAgreement(FlaskForm):
    accept = BooleanField("I have read and agree the terms and conditions.",  
            validators=[InputRequired(message="You must agree to the terms and conditions")])
    submit = SubmitField('                              Accept                               ', 
                         render_kw={"class": "btn btn-warning"})
    cancel = SubmitField('                              Cancel                               ',
                         render_kw={"class": "btn btn-danger"})

# Create Login Form
class LoginForm(FlaskForm):
	email = EmailField("Email :", validators=[DataRequired()])
	password = PasswordField("Password  :", validators=[DataRequired()])
	login = SubmitField("                            Login                              ")
	
class Forget_pswForm(FlaskForm):
    username = StringField("Username :", validators=[DataRequired()])
    email = EmailField("Email :", validators=[DataRequired()])
    send_reset = SubmitField('                   Send Reset Password                     ')
    
class Reset_pswForm(FlaskForm):
    password = PasswordField("Password  :" , validators=[DataRequired()])
    confirm_password = PasswordField("Confirm Password  :" , validators=[DataRequired()])
    reset = SubmitField('                          Reset Password                       ')
    
class SignupForm(FlaskForm):
    username = StringField("Username :", validators=[DataRequired()])
    email = EmailField("Email :", validators=[DataRequired()])
    password = PasswordField("Password  :" , validators=[DataRequired()])
    confirm_password = PasswordField("Confirm Password  :" , validators=[DataRequired(),
                                    EqualTo('password', 
                                    message='Passwords must match.')])
    # tick = BooleanField("I agree to the",  
    #         validators=[InputRequired(message="You must agree to the terms and conditions")])
    sign_up = SubmitField('                              Sign Up                               ')
    
class LogoutForm(FlaskForm):
    logout = SubmitField('  Logout  ')
    
class SupplierForm(FlaskForm):
    s_name = StringField("Name* :")
    s_tin = StringField('Tax Identification Number (TIN)* :', validators=[DataRequired()])
    s_business_no = StringField('Business Registration* :', validators=[DataRequired()])
    s_msic_code = StringField('MSIC Code* :', validators=[DataRequired()])
    s_msic_desc = TextAreaField('Business Activity Description* :',  
                                validators=[DataRequired(), Length(max=200)])
    s_telephone = TelField('Telephone Number* :', validators=[DataRequired()])
    s_sst = StringField("SST Registration Number** :", validators=[DataRequired()])
    s_email = EmailField('E-mail :')
    s_tt = StringField("Tourism Tax Registration Number** :", validators=[DataRequired()])
    s_add_1 = StringField("Address Line 1*", validators=[DataRequired()])
    s_add_2 = StringField("Address Line 2")
    s_add_3 = StringField("Address Line 3")
    s_city = StringField("City*", validators=[DataRequired()])
    s_postcode = StringField("Postcode* :", validators=[DataRequired()])
    s_state = StringField("State* :", validators=[DataRequired()])
    s_country = StringField("Country*", validators=[DataRequired()])
    
class BuyerForm(FlaskForm):
    b_id_type = StringField("ID Type* :", validators=[DataRequired()])
    b_id_no = StringField("Registration/Identification/Passport Number* :", 
                          validators=[DataRequired()])
    b_tel_no = TelField("Telephone Number* :", validators=[DataRequired()])
    b_email = EmailField('E-mail :')
    
    b_add_1 = StringField("Address Line 1*", validators=[DataRequired()])
    b_add_2 = StringField("Address Line 2")
    b_add_3 = StringField("Address Line 3")
    
    b_city = StringField("City*", validators=[DataRequired()])
    b_country = StringField("Country*", validators=[DataRequired()])
    
    b_tin = StringField("Tax Identification Number (TIN)* :", validators=[DataRequired()])
    b_name = StringField("Name* :", validators=[DataRequired()])
    b_sst = StringField("SST Registration Number** :", validators=[DataRequired()])
    b_tt = StringField("Tourism Tax Registration Number** :", validators=[DataRequired()])
    
    b_postcode = StringField("Postcode* :", validators=[DataRequired()])
    b_state = StringField("State* :", validators=[DataRequired()])
    
    b_search = SearchField("Search Alias / Buyer Name :",
                           render_kw={"style": "background-color: #90ee90;"})
    b_result = TextAreaField('Search Result :', validators=[Length(max=600)], 
                             render_kw={"rows": 4, "style": "background-color: #90ee90;"})
    
    def __init__(self, is_search=False, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Set optional validation if it's a search action
        if is_search:
            for field_name in ['b_id_type', 'b_id_no', 'b_tel_no', 'b_tin', 'b_name', 'b_sst', 
                               'b_tt', 'b_add_1', 'b_city', 'b_country', 'b_postcode', 'b_state']:
                getattr(self, field_name).validators = [Optional()]

# class DecimalInputWithComma(Input):
#     input_type = 'text'

#     def __call__(self, field, **kwargs):
#         if field.data:
#             kwargs['value'] = f"{field.data:,.2f}"
#         return super().__call__(field, **kwargs)
   
class DecimalInputWithComma(Input):
    input_type = 'text'

    def __call__(self, field, **kwargs):
        # Ensure that field.data is a real number (int or float)
        if field.data:
            try:
                # If field.data is not a number, convert it to float (if it's a string)
                numeric_value = float(field.data)  # Converts field.data to float
                kwargs['value'] = f"{numeric_value:,.2f}"  # Format with commas and two decimal places
            except (ValueError, TypeError):
                # If the value cannot be converted to float, retain the original value
                kwargs['value'] = field.data
        else:
            kwargs['value'] = ''  # If no data, leave the field empty
        return super().__call__(field, **kwargs)
    
    @property
    def validation_attrs(self):
        # Provide validation attributes as expected by WTForms
        return {}
                   
class LineItemsForm(FlaskForm):
    l_class_code = StringField("Classification Codes* :")
    l_desc = StringField("Description of Product or Service* :")
    l_quantity = IntegerField("Quantity* : ",
                              render_kw={"style": "background-color: #ffca2c;"})
    l_measurement = StringField("Measurement* :")
    l_toggle = BooleanField("Additional Information :")
    l_unit_price = DecimalField("Unit Price* (RM):", places=2, rounding=None,
                                # Format with comma
                                widget=DecimalInputWithComma())
    l_inv_no = StringField("Invoice Number*** :", validators=[DataRequired()],
                           render_kw={"style": "background-color: #ffca2c;"})
    l_date_time = DateTimeField("Date and Time Issued * :", validators=[DataRequired()],
                              
                                render_kw={"style": "background-color: #ffca2c;"})
    l_search = SearchField("Search Alias / Products or Services Name :",
                           render_kw={"style": "background-color: #90ee90;"})
    l_result = TextAreaField('Search Result :', validators=[Length(max=600)], 
                             render_kw={"rows": 4, "style": "background-color: #90ee90;"})
    
    def __init__(self, is_search=False, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Set optional validation if it's a search action
        if is_search:
            for field_name in ['l_toggle', 'l_class_code', 'l_desc', 'l_inv_no', 'l_date_time']:
                getattr(self, field_name).validators = [Optional()]
                
class AddInfoForm(FlaskForm):
    a_dis_desc =  SelectField("Discount Description :", validators=[Optional(),
                                                        Length(max=600)],
            choices=[('no','no'),
                ('Bulk Purchase Discount', 'Bulk Purchase Discount'), 
                        ('Early Settlement Discount', 'Early Settlement Discount'), 
                        ('Free Shipping', 'Free Shipping'),
                        ('Referral Discount', 'Referral Discount'), 
                        ('Upgrade Discount', 'Upgrade Discount'), 
                        ('other', 'Other')])
    a_dis_desc_custom = StringField("Custom Discount Description :",
            validators=[Optional(), Length(max=600)])
    a_dis_value = IntegerField("Discount Value(RM):", validators=[Optional()], default='0')
    a_fee_desc = SelectField("Fee / Charge Description :", validators=[Length(max=600)],
                            choices=[('no','no'),
                                     ('Delivery Charges', 'Delivery Charges'), 
                                    ('Installation Charges', 'Installation Charges'), 
                                    ('Packaging Charges', 'Packaging Charges'),
                                    ('other', 'Other')])
    a_fee_desc_custom = StringField("Custom Discount Description :", 
            validators=[Optional(), Length(max=600)])
    a_fee_value = IntegerField("Fee / Charge Value (RM) :", validators=[Optional()], 
                               default=0)
    a_payment_mode = SelectField("Payment Mode :",
                                 choices=[('NA','NA'),
                                        ('Bank Transfer', 'Bank Transfer'), 
                                        ('Cash', 'Cash'), 
                                        ('Cheque', 'Cheque'),
                                        ('Credit Card', 'Credit Card'), 
                                        ('Debit Card', 'Debit Card'),
                                        ('Digital Bank', 'Digital Bank'),
                                        ('E-wallet / Digital Wallet', 'E-wallet / Digital Wallet'),
                                        ('Others', 'Others')],)
    a_bank_acc = StringField("Supplier's Bank Account Number :")
    a_payment_term = StringField("Payment Terms :")
    a_payment_amount = DecimalField('Prepayment Amount :',widget=DecimalInputWithComma())
    a_prepayment_date = DateField("Prepayment Date :")
    a_prepayment_num = StringField("Prepayment Reference Number :")
    a_bill_num = StringField("Bill Reference Number :")
   
