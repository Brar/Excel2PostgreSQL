using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NpgsqlTypes;

namespace Excel2PostgreSQL
{
    internal class ColumnTypeInferer
    {
        public NpgsqlDbType ResultType { get; private set; }
        public bool ResultTypeMightBeDateOrTime { get; private set; } = true;

        private string _currentFormat;
        private dynamic _currentValue;

        private bool CurrentValueIsNumeric => _currentValue is double && ResultType != NpgsqlDbType.Text && ResultType != NpgsqlDbType.Boolean;
        private bool CurrentValueIsBool => _currentValue is bool && ResultType == NpgsqlDbType.Boolean;
        private bool CurrentValueIsPositiveOrZero => CurrentValueIsNumeric && _currentValue >= 0;
        private bool CurrentValueDoesNotHaveDecimalPlaces => CurrentValueIsNumeric && Math.Abs(_currentValue % 1d) < Double.Epsilon;

        private bool ResultTypeIsTimeStamCompatible => ResultType == 0 || ResultType == NpgsqlDbType.Date ||
                                                       ResultType == NpgsqlDbType.Time ||
                                                       ResultType == NpgsqlDbType.Timestamp;

        private bool CurrentValueIsMoney => CurrentValueIsNumeric && (ResultType == 0 || ResultType == NpgsqlDbType.Money) && _currentFormat.Contains('$');
        private bool CurrentValueMightBeDate => (_currentFormat.Contains('d') ||
                                     _currentFormat.Contains('y') ||
                                     _currentFormat.Contains('m')) &&
                                    CurrentValueDoesNotHaveDecimalPlaces &&
                                    CurrentValueIsPositiveOrZero &&
                                    (ResultType == 0 || ResultType == NpgsqlDbType.Date);

        private bool CurrentValueMightBeTime => (_currentFormat.Contains('h') ||
                                                 _currentFormat.Contains('s') ||
                                                 _currentFormat.Contains('m')) &&
                                                _currentValue >= 0 && _currentValue < 1 &&
                                                (ResultType == 0 || ResultType == NpgsqlDbType.Time);

        private bool CurrentValueMightBeDateTime => (_currentFormat.Contains('d') || _currentFormat.Contains('y') || _currentFormat.Contains('m')) &&
                                                    (_currentFormat.Contains('h') || _currentFormat.Contains('s') || _currentFormat.Contains('m')) &&
                                                    CurrentValueIsPositiveOrZero && ResultTypeIsTimeStamCompatible;


        public void UpdateType(dynamic value, string numberFormat)
        {
            if (ResultType == NpgsqlDbType.Text)
            {
                ResultTypeMightBeDateOrTime = false;
            }
            else
            {
                _currentValue = value;
                _currentFormat = numberFormat;
                ProcessCurrent();
            }
        }

        private void ProcessCurrent()
        {
            if (CurrentValueIsNumeric)
            {
                ProcessNumericValue();
            }
            else if (CurrentValueIsBool)
            {
                ResultTypeMightBeDateOrTime = false;
                ResultType = NpgsqlDbType.Boolean;
            }
            else
            {
                ResultTypeMightBeDateOrTime = false;
                ResultType = NpgsqlDbType.Text;
            }
        }

        private void ProcessNumericValue()
        {
            if (ResultType == NpgsqlDbType.Double)
            {
                ResultTypeMightBeDateOrTime = false;
            }
            else if (CurrentValueIsMoney)
            {
                ResultTypeMightBeDateOrTime = false;
                ResultType = NpgsqlDbType.Money;
            }
            else if (CurrentValueMightBeDate && !CurrentValueMightBeTime)
            {
                ResultType = NpgsqlDbType.Date;
                ResultTypeMightBeDateOrTime = false;
            }
            else if (CurrentValueMightBeTime && !CurrentValueMightBeDate)
            {
                ResultType = NpgsqlDbType.Time;
                ResultTypeMightBeDateOrTime = false;
            }
            else if (CurrentValueMightBeDateTime)
            {
                ResultTypeMightBeDateOrTime = false;
                ResultType = NpgsqlDbType.Timestamp;
            }
            else if (CurrentValueDoesNotHaveDecimalPlaces)
            {
                ResultTypeMightBeDateOrTime = false;
                ProcessInteger();
            }
            else
            {
                ResultType = NpgsqlDbType.Double;
            }
        }

        private void ProcessInteger()
        {
            if (ResultType == NpgsqlDbType.Bigint)
            {
                ResultType = NpgsqlDbType.Bigint;
                return;
            }
            if (ResultType == NpgsqlDbType.Integer)
            {
                ResultType = NpgsqlDbType.Integer;
                return;
            }

            double intValue = _currentValue > 0d ? Math.Floor(_currentValue) : Math.Ceiling(_currentValue);

            if (intValue < short.MaxValue || intValue > short.MinValue)
            {
                ResultType = NpgsqlDbType.Smallint;
            }
            else if (intValue < int.MaxValue || intValue > int.MinValue)
            {
                ResultType = NpgsqlDbType.Integer;
            }
            else
            {
                ResultType = NpgsqlDbType.Bigint;
            }
        }
    }
}
