
<# ==============================

Setting keyrate

                encoding: utf8bom
============================== #>

# https://github.com/cocolacre/scripts/blob/master/keyrate/keyrate.ps1

if(-not ('PersonalKeyRateSettings.KeyRate' -as [type]))
{Add-Type -Type @"

using System;
using System.Runtime.InteropServices;

namespace PersonalKeyRateSettings {
    public static class KeyRate {
        [DllImport("user32.dll", SetLastError = false)]
        internal static extern bool SystemParametersInfo(
              uint uiAction
            , uint uiParam
            , ref FILTERKEYS pvParam
            , uint fWinIni
        );

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct FILTERKEYS {
            public uint cbSize;
            public uint dwFlags;
            public uint iWaitMSec;
            public uint iDelayMSec;
            public uint iRepeatMSec;
            public uint iBounceMSec;
        }

        private const uint  SPI_SETFILTERKEYS   = 0x0033;
        private const uint  FKF_FILTERKEYSON    = 0x0001;
        private const uint  FKF_AVAILABLE       = 0x0002;

        private static void Usage() {
            Console.WriteLine("Usage: keyrate <delay ms> <repeat ms>");
        }

        public static void Main(string[] args) {
            if(args.Length != 0  && args.Length != 2) {
                Usage();
                Console.WriteLine("Call with no parameters to disable.");
                return;
            }

            FILTERKEYS fk;
            fk.cbSize       = 0;
            fk.dwFlags      = 0;
            fk.iWaitMSec    = 0;
            fk.iDelayMSec   = 0;
            fk.iRepeatMSec  = 0;
            fk.iBounceMSec  = 0;

            if(args.Length == 2) {
                fk.iDelayMSec   = (uint) Math.Max(0, Convert.ToInt32(args[0]));
                fk.iRepeatMSec  = (uint) Math.Max(0, Convert.ToInt32(args[1]));
                fk.dwFlags = FKF_FILTERKEYSON | FKF_AVAILABLE;
                Console.WriteLine("Setting keyrate: delay={0}ms, rate={1}ms"
                                  , fk.iDelayMSec, fk.iRepeatMSec);
            } else {
                Usage();
                Console.WriteLine("No parameters given: disabling.");
            }

            fk.cbSize = (uint) Marshal.SizeOf(fk);
            if(! SystemParametersInfo(SPI_SETFILTERKEYS, 0, ref fk, 0)) {
                Console.WriteLine("System call failed.\nUnable to set keyrate.");
            }
        }
    }
}
"@ }

function Set-Keyrate {
    param (
        [int]$delay
        ,[int]$rate
        ,[switch]$reset
    )
    $params = ($reset)? @() : @($delay, $rate)
    [PersonalKeyRateSettings.KeyRate]::Main($params)
}

Set-Keyrate -delay 180 -rate 12