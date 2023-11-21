import pandas as pd
import datetime as datetime
from dateutil.relativedelta import relativedelta

#
DATE_TYPES = (pd.Timestamp, datetime.datetime, datetime.date)


def is_valid_types(val=None, types=(), raise_on_invalid = False):
    '''
    Check whether a value falls within the allowed datatypes.
    
    Usage:
        > is_valid_types(1, (float, str)) => False
        > is_valid_types(1, (float, str), True) => TypeError
    '''
    
    if val is None:
        
        raise Exception ("Input should not be None.")
        
    elif not isinstance(val, types):
        
        if raise_on_invalid:
            
            # Prepare type error message
            
            if hasattr(types, '__iter__'):    
                types_str = ", ".join([t.__name__ for t in types])
            else:
                types_str = types.__name__ #for single value tuple
             
            raise TypeError (
                f"Input = {val} ({type(val).__name__}); "
                f"Expected the following types: ({types_str})."
                )
        
        return False
    
    else:
        
        return True


def convert_to_date(date = None):
    
    if date is None:
        
        raise Exception ("Date should not be None.")
    
    elif not isinstance(date, DATE_TYPES):
        
        raise Exception (f"Input date={date} is not date format: {type(date)}.")
    
    

class FYGenerator:
    '''
    The FYGenerator class generates a list of financial years with the
    start and end dates.

    It allows the generation of FY dates beyond the current FY, based
    on the number of prior and/or post periods required.

    The class also provides quick methods such as:
      - get_date_fy : that takes in a date and classify the FY
      - get_fy_dates: that takes in an offset and returns the FY start/end date
    '''

    def __init__(
        self,
        fy_start_date=None,
        fy_end_date=None,
        num_prior_periods=3,
        num_post_periods=3
    ):
        '''
        Construct a data structure to contain the FY for the given period.

        For example, FYGenerator('2020-03-01', None, 3, 3), will generate the 
        information for FY2017 ~ FY2023.

        Parameters
        ----------
        fy_start_date       : datetime.date;
        fy_end_date         : datetime.date;
        num_prior_periods   : int; (num_prior_periods > 0)
        num_post_periods    : int; (num_post_periods  > 0)

        Internal methods
        -------
        _create_fy_df    : to generate the dataframe that contains FY info

        Methods
        -------
        get_date_fy: 
            takes in a date and classify the FY

        get_fy_dates: 
            takes in an FY/offset and returns the FY start/end date
        '''

        # Save input as attributes
        self.input_parameters = {"num_prior_periods": num_prior_periods,
                                 "num_post_periods": num_post_periods,
                                 "fy_start_date": fy_start_date,
                                 "fy_end_date": fy_end_date}
        self.num_prior_periods = num_prior_periods
        self.num_post_periods = num_post_periods
        self.fy_start_date = fy_start_date
        self.fy_end_date = fy_end_date

        # Try to convert to datetime.date object, as long as it is of
        # a similar datatype (e.g. datetime.datetime, pd.Timestamp etc.)
        if fy_start_date is not None:

            self.fy_start_date = \
                self._convert_datetimestamp_to_date(self.fy_start_date)

        if fy_end_date is not None:

            self.fy_end_date =  \
                self._convert_datetimestamp_to_date(self.fy_end_date)

    def __repr__(self):

        # Get param
        input_parameters = self.input_parameters
        fy_start_date = input_parameters["fy_start_date"]
        fy_end_date = input_parameters["fy_end_date"]
        num_prior_periods = input_parameters["num_prior_periods"]
        num_post_periods = input_parameters["num_post_periods"]

        # Define the init str
        repr_list = (
            f"CLASS: FYGenerator(",
            f"           fy_start_date     = {fy_start_date},",
            f"           fy_end_date       = {fy_end_date},",
            f"           num_prior_periods = {num_prior_periods},",
            f"           num_post_periods  = {num_post_periods}",
            f"           )"
        )

        repr_str = '\n'.join(repr_list)

        # Print out df if available
        if hasattr(self, 'fy_df'):
            repr_str += "\n\nGenerated data:\n" + str(self.fy_df)

        # Add the key methods
        repr_str += "\n\nMethods:\n"
        repr_str += "  - get_date_fy(date)\n"
        repr_str += "  - get_fy_dates(offset)"

        return repr_str

    def _convert_datetimestamp_to_date(self, date_obj):
        '''
        Converts input of type pd.Timestamp, datetime.date, or 
        datetime.datetime to datetime.date format.        
        '''

        # Get the type
        date_obj_type = type(date_obj)

        # Valid types
        valid_types = [pd.Timestamp, datetime.date, datetime.datetime]

        # Convert to date
        if date_obj_type in valid_types:

            date = (
                date_obj.date()
                # only timestamp and datetime has date method
                if 'date' in dir(date_obj)
                else date_obj               # datetime.date has no date method
            )

        else:

            msg = (
                f'Invalid input type: {date_obj_type.__name__} {date_obj}.\n\n'
                f'Please ensure that input of the following types: '
                f'{[v.__name__ for v in valid_types]}.'
            )
            raise TypeError(msg)

        return date

    def _validate_parameters(self):
        '''
        This internal method performs the following:
            1) Check that num_prior_periods and num_post_periods are:
               (a) of integer type
               (b) non-negative
            2) Check that at least one of fy_start_date or fy_end_date 
               is provided
            3) Check that fy_start and end_dates are consistent if both
               are provided
            4) Calculates the start date if end date is provided, 
               or calculates the end date, if the start date is provided.
        '''

        # Get attributes
        num_prior_periods = self.num_prior_periods
        num_post_periods = self.num_post_periods
        fy_start_date = self.fy_start_date
        fy_end_date = self.fy_end_date

        # 1a)
        # Make sure num_prior_periods, num_post_period are of integer type
        prior_type = type(num_prior_periods)
        post_type = type(num_post_periods)
        if not ((prior_type is int) & (post_type is int)):
            msg = (
                f'Both num_prior_periods ({prior_type.__name__} {num_prior_periods}) '
                f'and num_post_periods ({post_type.__name__} {num_post_periods}) '
                f'must be integers.'
            )
            raise TypeError(msg)

        # 1b)
        # Make sure input parameter num_prior_periods, num_post_period
        #   are non-negative
        if not ((num_prior_periods >= 0) & (num_post_periods >= 0)):

            msg = (
                f'Both num_prior_periods ({num_prior_periods}) '
                f'and num_post_periods ({num_post_periods}) '
                f'must be non-negative.'
            )

            raise ValueError(msg)

        # 2)
        # If both fy_start_date, fy_end_date are None, raise error
        if (fy_start_date is None) & (fy_end_date is None):

            msg = (
                f'Either fy_start_date ({fy_start_date}) or '
                f'fy_end_date ({fy_end_date}) must be provided.'
            )

            raise ValueError(msg)

        # 3)
        # Special Case: 2020-02-29 to 2021-02-27
        # Check if both fy_start_date, fy_end_date are given, but not consistent
        elif (fy_start_date is not None) & (fy_end_date is not None):

            # Recompute the end date based on the start date
            recalculated_fy_end_date = \
                fy_start_date + \
                relativedelta(months=+12) + relativedelta(days=-1)

            # Raise error if cannot match
            if recalculated_fy_end_date != fy_end_date:
                msg = (
                    f'Inconsistent fy_start_date ({fy_start_date}) and '
                    f'fy_end_date ({fy_end_date}). '
                    f'Re-computed fy_end_date = {recalculated_fy_end_date}.'
                )

                raise ValueError(msg)

        # 4)
        # If only fy_start_date parameter is given,
        elif (fy_start_date is not None) & (fy_end_date is None):

            # If fy_start_date OR fy_end_date are parsed as timestamp, convert to datetime.date
            self.fy_start_date = self._convert_datetimestamp_to_date(
                fy_start_date)

            fy_end_date = fy_start_date + \
                relativedelta(months=+12) + relativedelta(days=-1)

            # Save as attr
            self.fy_end_date = fy_end_date

        # 4)
        # If only fy_end_date parameter is given,
        elif (fy_start_date is None) & (fy_end_date is not None):

            # If fy_start_date OR fy_end_date are parsed as timestamp, convert to datetime.date
            self.fy_end_date = self._convert_datetimestamp_to_date(fy_end_date)

            # Update fy_start_date
            fy_start_date = fy_end_date + \
                relativedelta(months=-12) + relativedelta(days=+1)

            # Adjusted so will not start on 29 Feb but 01 Mar instead
            # Else, in the case where end date is 28-02-2021, calculated start date will be 29-02-2020
            if (fy_start_date.day == 29 and fy_start_date.month == 2):
                fy_start_date = fy_start_date + relativedelta(days=+1)

            # save as attribute
            self.fy_start_date = fy_start_date

        # If both fy_start_date and fy_end_date parameters are given,
        else:

            # All okay at this stage.
            pass

        # End of function

    def _create_fy_df(self):
        '''
        Create a dataframe that contains the FY info of the given period.

        Output
        ------
        self.fy_df : dataframe
                     where index: prior to post periods
                         columns: [start, end, fy]
        '''

        # Generate the fy dates if it has not been generated.
        if not hasattr(self, 'fy_df'):

            # Validate the input parameters
            self._validate_parameters()

            # Create dataframe that contains FY start date and end date info
            #   for (prev_fy_n~future_fy_n) years
            num_prior_periods = self.num_prior_periods
            num_post_periods = self.num_post_periods

            self.fy_df = pd.DataFrame(
                index=pd.Index([], name='fy_offset'),
                columns=['start', 'end', 'fy', 'num_days']
            )

            # Define the range of periods
            fy_offsets = range(-self.num_prior_periods,
                               self.num_post_periods+1)

            # Calculate
            for fy_offset in fy_offsets:

                # Calculate the dates for the required offset
                self._calculate_fy_dates_by_offset_from_baseline(fy_offset)

        return self.fy_df

    def _calculate_fy_dates_by_offset_from_baseline(self, fy_offset):
        '''
        An internal method to calculate the fy dates for a given offset.

        Dates are calculated from baseline dates, i.e. fy_offset = 0.

        self.fy_df to be modified in-place.
        '''

        # Create the fy_df first, but i will assign to _ as we don't
        #   need the variable. we will just ref from self.
        _ = self._create_fy_df()

        # Calculate for fy_offset if it is not yet present in fy_df
        if fy_offset not in self.fy_df.index:

            # Calculate delta
            delta = relativedelta(years=fy_offset)

            # Calculate dates and save to df
            start = self.fy_start_date + delta
            end = self.fy_end_date + delta

            # for end date, adjust 28 feb to 29 feb whenever possible (leap year)
            end = self._adjust_for_leap_end(end)

            # if input fy_start_date is a 29 feb of a leap year,
            # without this block, the fy end for all the years including
            # leap years will be 27 feb.
            # Therefore, there will be a gap of one day to the next fy which
            # is a leap year.
            # For those affected years, will adjust 27 feb to 28 feb.
            # for start date (leap year), adjust previous years end date to 28 feb whenever possible
            if self._is_29feb_on_leap_year(self.fy_start_date):

                # One year before the leap year
                if (fy_offset % 4 == 3):

                    # Adjust end date from 27 feb to 28 feb
                    end = datetime.date(end.year, end.month, 28)

            # Calculate number of days in the fy
            num_days = (end - start).days + 1

            # Update the df
            self.fy_df.loc[fy_offset, ['start', 'end', 'fy', "num_days"]] = \
                [start, end, end.year, num_days]

        return self.fy_df.loc[fy_offset]

    def _is_29feb_on_leap_year(self, date):

        # Check if the start date is 29 feb on a leap year. if so, adjust the previous years end date to 28.
        if ((date.year % 4) == 0) and (date.month == 2) and (date.day == 29):

            return True

        else:

            return False

    def _adjust_for_leap_end(self, date):

        # Init output
        leap_adjusted_date = date

        # Check if the date is 28 feb on a leap year. if so, adjust to 29
        if ((date.year % 4) == 0) and (date.month == 2) and (date.day == 28):

            leap_adjusted_date = datetime.date(date.year, date.month, 29)

        return leap_adjusted_date

    def _fy_to_offset(self, fy):
        '''
        An internal method to convert fy to offset.
        '''

        # Get the base fy
        base_fy = self._create_fy_df().at[0, 'fy']

        return fy - base_fy

    def _offset_to_fy(self, offset):
        '''
        An internal method to convert offset to fy.
        '''

        # Get the base fy
        base_fy = self._create_fy_df().at[0, 'fy']

        return base_fy + offset

    def _verify_fy_offset(self, fy=None, fy_offset=None):
        '''
        Verify fy_offset based on the fy and fy_offset inputs.
        Returns
        -------
        verified_fy_offset
        '''

        # Initialise fy_offset_verified value as None, to be updated
        #   in the if-elif-else block.

        fy_offset_verified = None

        # no inputs are provided
        if (fy is None) and (fy_offset is None):

            msg = (
                f'Missing parameter(s): Either year ({fy}) or '
                f'offset ({fy_offset}) should be provided.'
            )

            raise ValueError(msg)

        # Check that if both year and fy_offset are provided, they are consistent
        elif (fy is not None) and (fy_offset is not None):

            recalculated_fy = self._offset_to_fy(fy_offset)

            # Inconsistent fy_offset and fy - raise Error
            if recalculated_fy != fy:
                msg = (
                    f'Inconsistent year ({fy}) and '
                    f'fy_offset ({fy_offset}). '
                    f'Re-computed year = {recalculated_fy}.'
                )

                raise ValueError(msg)

            else:

                # If they are consistent, then will use the fy_offset
                fy_offset_verified = fy_offset

        elif fy is not None:

            # fy_offset not provided but fy is provided. so calculate offset.
            fy_offset_verified = self._fy_to_offset(fy)

        elif fy_offset is not None:

            # fy not provided, but fy_offset is provided. Use that.
            fy_offset_verified = fy_offset

        else:

            # If it reaches this block, something is wrong because
            #   the 4 scenarios above should cover all the cases.

            raise NotImplementedError("Unexpected error")

        return fy_offset_verified

    def get_fy_dates(self, fy=None, fy_offset=None, which='both'):
        '''
        Get FY start date AND/OR FY end date of year of interest
        Parameters
        ----------
        fy          : int; fy of interest
        fy_offset   : int;
        which       : string; options = ['both', 'start', 'end']
                      If both, tuple will be returned

        Returns
        -------
        datetime.date   (if which=='start'|which=='end')
        tuple           (if which=='both')

        Examples
        --------
        >>> end_date = pd.to_datetime('2020-03-31')
        >>> FY = FYGenerator(fy_end_date=end_date)
        >>> p3fy = FY.get_fy_dates(fy_offset=5, which='both')
        >>> print(p3fy)
           (datetime.date(2025, 4, 1, 0, 0), datetime.date(2026, 3, 31, 0, 0))
        '''

        # Get verified fy_offset
        fy_offset_verified = self._verify_fy_offset(fy=fy, fy_offset=fy_offset)

        # Create or get fy_df
        fy_df = self._create_fy_df()

        # Calculate dates for this offset from baseline dates
        data = self._calculate_fy_dates_by_offset_from_baseline(
            fy_offset_verified)

        # Get the start and end date
        start_date, end_date = data.loc[['start', 'end']]

        # If invalid parameter name is given for 'which'
        which_options = ['both', 'start', 'end']
        if which not in which_options:
            msg = (
                f'Invalid parameter for which = {which}.\n\n'
                f"Please set 'which' to either one of {which_options}."
            )
            raise TypeError(msg)

        elif which == 'both':

            return (start_date, end_date)

        elif which == 'start':

            return start_date

        else:

            # since we have verified the options in the first condition,
            # if it reaches this block, which = 'end'

            return end_date

    def get_date_fy(self, date):
        '''
        Get the FY of the given input date.

        Parameter
        ---------
        date: datetime.date;

        Returns
        -------
        int; e.g. 2019

        Example
        -------
        >>> end_date = pd.to_datetime('2020-03-31')
        >>> FY = FYGenerator(fy_end_date=end_date, num_prior_periods=3, num_post_periods=3)
        >>> print(FY.get_date_fy(pd.to_datetime('2019-03-31')))
            2019
        '''

        # Convert user input date to date object
        date = self._convert_datetimestamp_to_date(date)

        # Get a reference start and end date based on the input year
        date_fy = None
        for offset in [0, 1]:

            # Get the ref fy_start and end
            fy = date.year+offset
            fy_start, fy_end = self.get_fy_dates(fy=fy, which='both')

            # Check if date is within this range
            if (fy_start <= date) and (date <= fy_end):
                date_fy = fy
                break

        # Return
        if date_fy is None:
            raise NotImplementedError("Unexpected error.")
        else:
            return date_fy

if __name__ == "__main__":
    
    # Tester for FYGenerator
    if True:

        # 1) Provide start date
        if True:
            start_date = datetime.date(2020, 7, 1)
            end_date = None
            num_prior_periods = 0
            num_post_periods = 3
            self = FYGenerator(fy_start_date=start_date,
                               fy_end_date=end_date,
                               num_prior_periods=num_prior_periods,
                               num_post_periods=num_post_periods)

            print(self.get_date_fy(datetime.date(2025, 1, 30)))

        # 2) Provide start date that does not start on 1st day of month
        #   Note that the leap year is treated correctly. For non-leap
        #   years, the start dates are adjusted to 28 feb.

        # Start on 29
        if False:
            start_date = datetime.date(2020, 2, 29)
            end_date = None
            num_prior_periods = 5
            num_post_periods = 5
            self = FYGenerator(fy_start_date=start_date,
                               fy_end_date=end_date,
                               num_prior_periods=num_prior_periods,
                               num_post_periods=num_post_periods)
            self._create_fy_df()
            print(self.fy_df)
            print(self.get_date_fy(datetime.date(2025, 2, 28)))

        # Start on 28
        if False:
            start_date = datetime.date(2020, 2, 28)
            num_prior_periods = 5
            num_post_periods = 5
            self = FYGenerator(fy_start_date=start_date,
                               num_prior_periods=num_prior_periods,
                               num_post_periods=num_post_periods)
            self._create_fy_df()
            print(self.fy_df)
            print(self.get_date_fy(datetime.date(2025, 2, 28)))

        # 3) Provide end date that crosses 29 feb.
        #   Note that the leap year adjustment to 29 feb will take place
        #   automatically.
        # End on 28
        if False:
            start_date = None
            end_date = datetime.date(2020, 2, 28)
            num_prior_periods = 5
            num_post_periods = 5
            self = FYGenerator(fy_start_date=start_date,
                               fy_end_date=end_date,
                               num_prior_periods=num_prior_periods,
                               num_post_periods=num_post_periods)
            self._create_fy_df()
            print(self.fy_df)
            print(self.get_date_fy(datetime.date(2025, 2, 28)))

        # End on 29
        if False:
            end_date = datetime.date(2020, 2, 29)
            num_prior_periods = 5
            num_post_periods = 5
            self = FYGenerator(fy_end_date=end_date,
                               num_prior_periods=num_prior_periods,
                               num_post_periods=num_post_periods)
            self._create_fy_df()
            print(self.fy_df)
            print(self.get_date_fy(datetime.date(2025, 2, 28)))

        # 4) Error: Inconsistent start and end dates provided.
        if False:
            start_date = datetime.date(2019, 3, 1)
            end_date = datetime.date(2019, 3, 2)
            num_prior_periods = 0
            num_post_periods = 3
            self = FYGenerator(fy_start_date=start_date,
                               fy_end_date=end_date,
                               num_prior_periods=num_prior_periods,
                               num_post_periods=num_post_periods)

            print(self.get_date_fy(datetime.date(2025, 2, 28)))

        # 5) Error: Negative periods provided.
        if False:
            start_date = datetime.date(2020, 2, 29)
            num_prior_periods = -5
            num_post_periods = 5
            self = FYGenerator(fy_start_date=start_date,
                               num_prior_periods=num_prior_periods,
                               num_post_periods=num_post_periods)
            self._create_fy_df()
            print(self.fy_df)
            print(self.get_date_fy(datetime.date(2025, 2, 28)))