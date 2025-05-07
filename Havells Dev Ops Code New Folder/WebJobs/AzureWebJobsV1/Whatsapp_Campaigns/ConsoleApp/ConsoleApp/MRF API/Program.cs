using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MRF_API
{
    internal class Program
    {
        static IOrganizationService _service;
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            _service = HavellsConnection.CreateConnection.createConnection(finalString);


            MRFAPI obj = new MRFAPI(_service);

            //string abc = "{\"SalesTrackingID\ =4773956,\"InvoiceDate\ =\"2022-09-23T00:00:00\",\"InvoiceNo\ =\"59780855\",\"InvoiceValue\ =25178.0,\"PurchaseLocationPINCode\ =500003,\"Attachment\ =\"PDFconvertedinBase64Format\",\"FileName\ =\"1ba9f4fa-0743-4ecb-92f5-9fa46a0f3732_SalesTracking.jpg\",\"PurchaseFromCode\ =\"CAE0943\",\"PurchaseFromName\ =\"SRIKALYANIHOUSEOFELECTRONICS\",\"Source\ =\"19\",\"Address1\ =\"HOUSENO8-1\",\"Address2\ =\"3523FLOORNEARVENKATESHSWMAYTEMPLE\",\"Address3\ =\"VivekVihar\",\"CustomerMobileNo\ =\"8285906486\",\"CustomerFirstName\ =\"Kuldeep\",\"CustomerLastName\ =\"Khare\",\"PINCode\ =\"201304\",\"ServiceCallData\ =[{\"CallType\ =\"I\",//I-forInstallationandDemo&PfornoInstallation.\"SerialNumber\ =\"9874123655\",\"ModelNumber\ =\"GHWRDLK015\",\"Qty\ =1,\"ServiceDate\ =\"2022-12-22T16:03:12.203\"}]}";


            List<SFA_ServiceCallData> lst = new List<SFA_ServiceCallData>();

            lst.Add(new SFA_ServiceCallData
            {
                CallType = "I",
                SerialNumber = "9874123655",
                ModelNumber = "GHWRDLK015",
                Qty = 1,
                ServiceDate = "2022-12-22T16:03:12.203"
            });
            SFA_ServiceCall serviceCallData = new SFA_ServiceCall()
            {
                SalesTrackingID = "9876545678787656786435674",
                InvoiceDate = "2022-09-23T00:00:00",
                InvoiceNo = "000000000000",
                InvoiceValue = 765.34M,
                PurchaseLocationPINCode = "000000",
                Attachment = "iVBORw0KGgoAAAANSUhEUgAAALQAAAAwCAYAAAC47FD8AAAACXBIWXMAAAsTAAALEwEAmpwYAAAKTWlDQ1BQaG90b3Nob3AgSUNDIHByb2ZpbGUAAHjanVN3WJP3Fj7f92" +
                "UPVkLY8LGXbIEAIiOsCMgQWaIQkgBhhBASQMWFiApWFBURnEhVxILVCkidiOKgKLhnQYqIWotVXDjuH9yntX167+3t+9f7vOec5/zOec8PgBESJpHmomoAOVKFPDrYH49PSMTJvYACFUjgBCAQ" +
                "5svCZwXFAADwA3l4fnSwP/wBr28AAgBw1S4kEsfh/4O6UCZXACCRAOAiEucLAZBSAMguVMgUAMgYALBTs2QKAJQAAGx5fEIiAKoNAOz0ST4FANipk9wXANiiHKkIAI0BAJkoRyQCQLsAYFWBUiwC" +
                "wMIAoKxAIi4EwK4BgFm2MkcCgL0FAHaOWJAPQGAAgJlCLMwAIDgCAEMeE80DIEwDoDDSv+CpX3CFuEgBAMDLlc2XS9IzFLiV0Bp38vDg4iHiwmyxQmEXKRBmCeQinJebIxNI5wNMzgwAABr50cH+OD+Q5+bk4" +
                "eZm52zv9MWi/mvwbyI+IfHf/ryMAgQAEE7P79pf5eXWA3DHAbB1v2upWwDaVgBo3/ldM9sJoFoK0Hr5i3k4/EAenqFQyDwdHAoLC+0lYqG9MOOLPv8z4W/gi372/EAe/tt68ABxmkCZrcCjg/1xY" +
                "W52rlKO58sEQjFu9+cj/seFf/2OKdHiNLFcLBWK8ViJuFAiTcd5uVKRRCHJleIS6X8y8R+W/QmTdw0ArIZPwE62B7XLbMB+7gECiw5Y0nYAQH7zLYwaC5EAEGc0Mnn3AACTv/mPQCsBAM2XpOMAALzoG" +
                "FyolBdMxggAAESggSqwQQcMwRSswA6cwR28wBcCYQZEQAwkwDwQQgbkgBwKoRiWQRlUwDrYBLWwAxqgEZrhELTBMTgN5+ASXIHrcBcGYBiewhi8hgkEQcgIE2EhOogRYo7YIs4IF5mOBCJhSDSSgKQg6" +
                "YgUUSLFyHKkAqlCapFdSCPyLXIUOY1cQPqQ28ggMor8irxHMZSBslED1AJ1QLmoHxqKxqBz0XQ0D12AlqJr0Rq0Hj2AtqKn0UvodXQAfYqOY4DRMQ5mjNlhXIyHRWCJWBomxxZj5Vg1Vo81Yx1YN3YVG" +
                "8CeYe8IJAKLgBPsCF6EEMJsgpCQR1hMWEOoJewjtBK6CFcJg4Qxwicik6hPtCV6EvnEeGI6sZBYRqwm7iEeIZ4lXicOE1+TSCQOyZLkTgohJZAySQtJa0jbSC2kU6Q+0hBpnEwm65Btyd7kCLKArCCX" +
                "kbeQD5BPkvvJw+S3FDrFiOJMCaIkUqSUEko1ZT/lBKWfMkKZoKpRzame1AiqiDqfWkltoHZQL1OHqRM0dZolzZsWQ8ukLaPV0JppZ2n3aC/pdLoJ3YMeRZfQl9Jr6Afp5+mD9HcMDYYNg8dIYigZaxl7" +
                "GacYtxkvmUymBdOXmchUMNcyG5lnmA+Yb1VYKvYqfBWRyhKVOpVWlX6V56pUVXNVP9V5qgtUq1UPq15WfaZGVbNQ46kJ1Bar1akdVbupNq7OUndSj1DPUV+jvl/9gvpjDbKGhUaghkijVGO3xhmNIRb" +
                "GMmXxWELWclYD6yxrmE1iW7L57Ex2Bfsbdi97TFNDc6pmrGaRZp3mcc0BDsax4PA52ZxKziHODc57LQMtPy2x1mqtZq1+rTfaetq+2mLtcu0W7eva73VwnUCdLJ31Om0693UJuja6UbqFutt1z+o+02P" +
                "reekJ9cr1Dund0Uf1bfSj9Rfq79bv0R83MDQINpAZbDE4Y/DMkGPoa5hpuNHwhOGoEctoupHEaKPRSaMnuCbuh2fjNXgXPmasbxxirDTeZdxrPGFiaTLbpMSkxeS+Kc2Ua5pmutG003TMzMgs3KzYrMn" +
                "sjjnVnGueYb7ZvNv8jYWlRZzFSos2i8eW2pZ8ywWWTZb3rJhWPlZ5VvVW16xJ1lzrLOtt1ldsUBtXmwybOpvLtqitm63Edptt3xTiFI8p0in1U27aMez87ArsmuwG7Tn2YfYl9m32zx3MHBId1jt0O3x" +
                "ydHXMdmxwvOuk4TTDqcSpw+lXZxtnoXOd8zUXpkuQyxKXdpcXU22niqdun3rLleUa7rrStdP1o5u7m9yt2W3U3cw9xX2r+00umxvJXcM970H08PdY4nHM452nm6fC85DnL152Xlle+70eT7OcJp7WMG3" +
                "I28Rb4L3Le2A6Pj1l+s7pAz7GPgKfep+Hvqa+It89viN+1n6Zfgf8nvs7+sv9j/i/4XnyFvFOBWABwQHlAb2BGoGzA2sDHwSZBKUHNQWNBbsGLww+FUIMCQ1ZH3KTb8AX8hv5YzPcZyya0RXKCJ0VWhv6" +
                "MMwmTB7WEY6GzwjfEH5vpvlM6cy2CIjgR2yIuB9pGZkX+X0UKSoyqi7qUbRTdHF09yzWrORZ+2e9jvGPqYy5O9tqtnJ2Z6xqbFJsY+ybuIC4qriBeIf4RfGXEnQTJAntieTE2MQ9ieNzAudsmjOc5JpU" +
                "lnRjruXcorkX5unOy553PFk1WZB8OIWYEpeyP+WDIEJQLxhP5aduTR0T8oSbhU9FvqKNolGxt7hKPJLmnVaV9jjdO31D+miGT0Z1xjMJT1IreZEZkrkj801WRNberM/ZcdktOZSclJyjUg1plrQr1zC3" +
                "KLdPZisrkw3keeZtyhuTh8r35CP5c/PbFWyFTNGjtFKuUA4WTC+oK3hbGFt4uEi9SFrUM99m/ur5IwuCFny9kLBQuLCz2Lh4WfHgIr9FuxYji1MXdy4xXVK6ZHhp8NJ9y2jLspb9UOJYUlXyannc8o5" +
                "Sg9KlpUMrglc0lamUycturvRauWMVYZVkVe9ql9VbVn8qF5VfrHCsqK74sEa45uJXTl/VfPV5bdra3kq3yu3rSOuk626s91m/r0q9akHV0IbwDa0b8Y3lG19tSt50oXpq9Y7NtM3KzQM1YTXtW8y2rNv" +
                "yoTaj9nqdf13LVv2tq7e+2Sba1r/dd3vzDoMdFTve75TsvLUreFdrvUV99W7S7oLdjxpiG7q/5n7duEd3T8Wej3ulewf2Re/ranRvbNyvv7+yCW1SNo0eSDpw5ZuAb9qb7Zp3tXBaKg7CQeXBJ9+mfH" +
                "vjUOihzsPcw83fmX+39QjrSHkr0jq/dawto22gPaG97+iMo50dXh1Hvrf/fu8x42N1xzWPV56gnSg98fnkgpPjp2Snnp1OPz3Umdx590z8mWtdUV29Z0PPnj8XdO5Mt1/3yfPe549d8Lxw9CL3Ytslt" +
                "0utPa49R35w/eFIr1tv62X3y+1XPK509E3rO9Hv03/6asDVc9f41y5dn3m978bsG7duJt0cuCW69fh29u0XdwruTNxdeo94r/y+2v3qB/oP6n+0/rFlwG3g+GDAYM/DWQ/vDgmHnv6U/9OH4dJHzEfVI" +
                "0YjjY+dHx8bDRq98mTOk+GnsqcTz8p+Vv9563Or59/94vtLz1j82PAL+YvPv655qfNy76uprzrHI8cfvM55PfGm/K3O233vuO+638e9H5ko/ED+UPPR+mPHp9BP9z7nfP78L/eE8/sl0p8zAAAAIGNIU" +
                "k0AAHolAACAgwAA+f8AAIDpAAB1MAAA6mAAADqYAAAXb5JfxUYAAA37SURBVHja7J17kBzVdcZ/t3tmd2ZXu1o9Vm8JIfTgIRSCpBBhHLAFVqI42CEmiU3FsU3KsSsmiV0pknLlHxKX8yibqsQulxMqIU" +
                "mZ2GUHmxDKcoGNwRYQbIxAIEDItiQWSQi00u5qta+Z6Zs/7tfZu709szOzK2lt+qvqktTTfafvud8995zvnh4Zay0ZMvy8IMhMkCEjdIYMGaEzZMgInSFDRugMb16YnlVrJ56JLJQrtGxbTbi8C4xx59" +
                "JuLuQZe/YIpb1HMIV82iXrgI8Ay4F/Bx4CypnZM5wt5Oq4plXXDQNRg+1vBt4FrALygAGeBE4CmV6Y4ZyHHAbYAdwKXNRE+9fqvjxwE/AN4LPA6sz0Gc4Hod8vAv418DdAYRLdKxG2HLm/p3t3Hy3A7w" +
                "D3iuwZMpwDQluLHa2stnALgVkLdALXAL834bLRMvn1i2i5Yjl2tJzWdmtK6wXgF4F/Bt6ZDUGGs0doF9UWac1fbdpb3mcCs9FLCBcCt0zw0pWIYEE74eJOqERpbedrfPd6ef13ZMOQ4WwRuoi1twbzi5" +
                "uDruKvYswCj9AhsAG4HAixFtPWQuVoP+WXX4d8mGy7AyhO8f0bgT8HNmVDkWGmVY482G3AbWH3nC/QkluhmNdHEbiOwOwzYTgU9Q4xtudVKj2nMB0FmFjo9EvAhXUmjn8M/Ckw2MBEXAF0ASUvojfAa0" +
                "AfUPGuD4F5QLeusfozAPqBV6dQXdq8RNZ6bQ4Bx4EzievnAIt1XznxHIdq9HO17q14/awAJ4Be3b8MmEvj8mcOpy69AYzp3wt04Nkk1HcdrbNdA6zUc8ftxM8df18zitYq9bVL/S1qrPvV7jHgYC1CXw" +
                "h8jMB0BnOL1gQUmFyJVwS2mZbcXUFbfmjs8YOUXzlF0DmJzAsUb2+o48FDYCfwIPDVOjvbJuXlLSIVHkHvBr6VIE1Roc3NCoMiXZ8HfgjcUYMgBeCtwMcShG7V5LkXuD9xzxLgQ8BWnNzpJ8V3At9msg" +
                "S6EPiEQrExb3wGgS8Cj2pwPwxclWi3HhRl43tE1g6cpHqj+hQTuhX4DvB3DRD6I1q5Y/vEk32Xvm+0gedcILu9G7gaWAO0e5+PAT/Fyb//DTwupzKB0CGwDfhN4HkiG4FxurPByPiRe1izxg6P5aIogjD" +
                "AtIQ+mVtkuE8o4cvV2YmlwB/pAevpfKtI9raUz54EHk5RVzbJSEl0AZ+qQehuqT3vrDG5HkgQtF/e5fqU6/cCP9Aq4q84W4H3yBY+eoG/lcfr0Ip2TZMr8mngf0ToNmAL8Osp140Bf1+nZ42l3SsT50" +
                "vyzl9p4PkWa7X+ixoKXAtwsY7fBz4D/GO8ysaEuwDYLuquKB8Z6GtZ1PGImdN6gy3bItaeILKnsLYfGx0a+fb+YTtUgtBgWnM+Ma4BPiqiFRs09joZZpcXRlRDVIP4Iynez8qjjaWEUX1TDNxC4Feme" +
                "O6lIknczgngCRk8iQ1aovsSpNic8mxlkX9vnIZ73rsZDHsTt5ISKsUYaDBMOFVlHAYbaCMA/gS4ncZKMj6u1eCzwEDMxsuAHRgDIZ2VYwOfHCtVvmzmtd0BtoWRUtGORW22EhWpVAp2pPRprG3Hmjka" +
                "hDhmXCJitzRh7C55qAfrIPS5gpFtltW4phu4DviyRwILHBARN6UkwsuA5xLfs0UeOEmUR72Y2kyzP+0NrJrnGu9WhBA2eF8I/CHwFPBATsnVFcCi/58pUXRJ5Y3BD9B7ZsCAgShvI5PD2jyQJzAdhCY" +
                "v4wQzZKSCPOFi4PA0SdjI+VpYrtCmlseYD/wa8J+J8y8qvksSeiUTd0qN+rwxxRGcBL5Zp7fcCxyRB08jRR7YLe87G3E1sDbl/H3KUY4Alyo3uzIlVHkH8GpOS90lE52/gciuoRQ5SwbTpEX9RFwij3" +
                "h0Gl76VGI5j8/1N9HWGqbWyQvKP5bL6DH5jgP/q4QpSaz1um9E/96o0CaJ/ZoY9eAeXPHXaJVRyikePzFLCb20imPcrUQfrd6vKEHPJ3KqU0BnTkvdZHktMO44twhxu4iPp5CyXvyCsvdBT6Kbo4nSa" +
                "LnsRibXnfxU4dH8hKpxHfBfImkc57+kAViVaGO98pb9IvbVTN6E6gUea0CeewHY8zMsIVcrfNshLjwte35d4VqbVrS8CP06cDIn4y6dRR1bR7JmpDHcDGyXhw9kqLwIGDboMa5KOf8lhWg3JrzfuyQX" +
                "jnjnj8hr3prSx4tF6KLkx3zKxHm4gefdqUEeTGkrUHJ5lNlb5XhUkzfppW/QhN8tifU5Td6jUlEmLUMrEt7mfCdhS6i9ZT4VunVMF1twm0NJ1eEBxaHbPX00LzltkTxrTJrXRcoPJlaHVV682KU4Ozm" +
                "QB4BnGnjej+qohvfidP7ZSujHJCFeViWZ3aED3KbK43Ig31fONRLP3IUp2fX5JPScJjLds/Ecl8uL+jgsz7AXp3v613fjtGS/IGsMeFbE9olUFKkDxendibj3jMg8ky9D2LOaAU0f9+M2qUbrmHRLgd" +
                "8C7lKYt1NOxcyUQjGTCGaB4ecpfk4S4ikZfD/w45T7tntqkZ+Q7k4h5xKRen1KOy+ehXjYMvtfqvgH3K5tI/r1pcC/4grnOgKcKD1bXouyIkx0np9ja8rSZ3Ca8GnFxi+k3He9JCQf/Yqjk1iE24jaVCXB+" +
                "+EM9ynP7MdJ4Auy418prJhqIynA1Xq8B7glh6tHWMp4kcr5JvSJaU6w+2WIfg1iGVfPfa0kuHrCmbekeM7TnocG2KdzHYmlcJO8a9yHYU2Ekwmyr1YCuyYl299D4zLjl3C7k8Mpq66RTX4WXnvrVwK7ByfX" +
                "rcPJypvkAKrVB10F5HK46q8Vs4TQcZw6Oo37v6/YyidEl8i9vQ5Cz8VJh0mlJcRVBPZpVbusSlvX4op7DnkEfQX4kTxPvHmyTMl48iWIl4Dnm+j3N3G7lTOB8gw5pzL1bdXP1wSfK3XKKskb0Or2EE7m3CBn" +
                "814R3cdCYENOnma1nwApiB1S/Jcz0G7dAOfiETqLU/05xivomkElZUDKDYQx20jfsWoD3lfH/W+TLQ9550q4GpVf9hSlFtJLBHY3Sei4MGyYdL098Owz1fCFcgBpiaRfrDaVTXOy2xmq7wGUlE/cJicx5n3HM" +
                "PBJrYyHdTwiYq9Lm0A5uff18h7xTsTLw8buLcMp4+rp2g0UDbQaS2sALcH4lneoPzsUF7ZOg4yntdSMTDNWzNdxrhp24Hb9msUKDcxDjNdgRLhdrtuoLZFWNHivNfG9O6WWDFdZOYySre/iNO5a2CAJMI3QgRd+PT" +
                "vFBNuMK3cdqUJoi9vhPC2Cbq3CibuAHj3LJaTX1hwHXsjJI+4DygZCA28MGHvna0FUGIELQgwhHAjgTAhjgaVvvjWPdlhjPc/dqUlxgzxUdxPSW1ly2MHzmBS24rTnjhSiDaUMRk6hSXKwtsjoPR6hX5Y6sraGt+qpkmzW" +
                "g9/WUQtDwAfqIPRmHbXauWMKQheAt+uohWeB72nCb04J9W5SqPiS7H0xE+ujYzwDPJjTDHge2G9gQwkeOhZE84YNHwcWWSzlcc3HGsOTXRWzy0wuGXwM+DdcMc+duAKSRuS3AdymxfmotIsL/t+a4p0j4Ce4LVe8Po0qrrs+5Z4" +
                "rNTg9ifNPKOyotvHzXep/U6QZ9DG98lM/cRuZgXZ8zfke5S43pkz4uaTv2voO51vA1+Js+BkDD1q4bMxwuGTYCrSl1CSNAC/3hFGlvRJQtCbJPqts+v24X0ra0kDneoGv1WlwU2MFCKvEfbkqnrGV8deGfoPJRUJxUfynUjz0" +
                "ShnzQ4nP1qrv9yXOPyK9tBqhn2Dihk0j/a43HAsTNmkGoXevpflShVZvvA7iXmTokEduBJ+WynMy5824h4GbA8tiA6MGKikB1CiwJ4LSgSDiwkpAJyZZ3lXWsnm3ZKqVdXrn++TR6gk3rOchRrwBMpoQNsXLjii+LIiE8etGp/V5h0" +
                "gYep6jII/5HdKL4X+iJO6DCWWmwHiNzDHv/NMK7y7VShSNpy0cU1xammIlGU70u5GkcdBrP7bhGI1tuhQUcvh2P814lV+94WKgNspe336Ee2PlJoVGtX7caFCK1n8oX+mFib9tt9DAhyuw82AYfW/I2D8Aus3kwPtG4OkylIvWsDI" +
                "KmGfdkyWuXaYQ5IY6OvcE8LuSt+r1NFcoARtNeLAXNTFKiesvkOcMvcHLibBPidzXaMD8z/sUn1WrI16sZCZKeJ7jCuWS912uZ0mSKE60ar0rWJQeu6wJaTPeRNunZ2uVBr6mQdEqfl/wx1IdjPq/0Juc9cAw/kZOMnyNd2rXyVaLNC6j" +
                "jL8ge0Dj/KJv+wk/1mhcGentRwO7vzeIbo9gaTDRK/5AQf5QrLe0W8Mya5gbmQmvWQt3a6bVQg/wZ9T/gmyGNxcKjL8FVdLEH6ylE/o4DPzTHOjqgyOR0weN50F2+dl+C9BvLDkD89MJPZWe3Ivb6szInKEaRhqRMZOEjgw8U7CY0HJvyXCR3D+STb6aDOjagc6oauBUa1k8BfwL7ifBMmSYEQRVMi" +
                "4bwd3W/STAMK788T4SrwOVgUVRwPIoqLZXOlxD9rkL96r8yWwYMswUask2xyVVHcJtCny+WtpdI62NX5v3v+cE8Je4uoOBbAgynCtCoyRwn4LxZjY8XlWcvFj3fwX4HE7WO5OZP8O5JnSJ5l9WBacPXoTbOduF+2WknszsGc4WTPY/yWb4uU8KM2TICJ0hQ0boDBkyQmfIkBE6w5sD/zcAPQ/YokvKLw4AAAAASUVORK5CYII=",
                FileName = "1ba9f4fa-0743-4ecb-92f5-9fa46a0f3732_SalesTracking.jpg",
                PurchaseFromCode = "",
                PurchaseFromName = "SRI KALYANI HOUSE OF ELECTRONICS",
                Source = "19",
                Address1 = "HOUSE NO 8-1",
                Address2 = "352 3 FLOOR NEAR VENKATESH SWMAY TEMPLE ",
                Address3 = "Vivek Vihar",
                CustomerMobileNo = "8285906486",
                CustomerFirstName = "Kuldeep",
                CustomerLastName = "Khare",
                PINCode = "201304",
                ServiceCallData = lst
            };




            obj.SFA_CreateServiceCall(serviceCallData);
        }
    }
}
