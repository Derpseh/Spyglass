#!/usr/bin/env python2

# UpdTime -- Scrapes NationStates API to estimate current update lengths.
# Copyright (C) 2017   Khronion <khronion@gmail.com>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
# IMPORTANT: The above applies SOLELY to UpdTime, and does not apply to any
# program or script it may come bundled with.

from __future__ import with_statement
from __future__ import absolute_import
import urllib2
import datetime
import xml.etree.ElementTree as et


class UTC(datetime.tzinfo):
    # code derived from example code in Python2 documentation:
    # https://docs.python.org/2/library/datetime.html#tzinfo-objects
    u"""UTC"""

    def utcoffset(self, dt):
        return datetime.timedelta(0)

    def tzname(self, dt):
        return u"UTC"

    def dst(self, dt):
        return datetime.timedelta(0)


class UpdTime:
    def __init__(self, user_agent):
        self.user_agent = user_agent

    @staticmethod
    def timestamp(dt):
        return (dt - datetime.datetime(1970, 1, 1).replace(tzinfo=UTC())).total_seconds()

    def get(self):

        headers = {u'User-Agent': self.user_agent}

        # list of the last couple of late updaters in the game
        late_updaters = u"Violet_Irises,Hiseville,The_First_Council_of_Nationstates,DSDZ_Warzone," \
                        u"The_United_Regions_of_Krachting,The_Pariahan_Imperium,The_Islands_of_Somewhere,Etoron," \
                        u"Cayenne,Sweetcorn,Kodiak,The_Twelve_Majestic_Isles,Fayu,Gomu,Welnat,Cusco,Dolph,Dat_Boi," \
                        u"Sicherungskopie_bunicken,The_United_Federal_Agencian_Republic,Simsbury," \
                        u"The_Holy_Rex_Strikepire,The_InternetZ,Hertzia,Hampshire,Kami_No_Chikara,Whirlpool_Galaxy," \
                        u"Old_Celtica,Radicali_Italiani,The_Yonderlands,By_the_Sword,Jackson_Heights,Suwoville,Oscar," \
                        u"The_Empire_of_Agarian,Versailles_Isle,Golden_Bow_of_Apollo," \
                        u"The_Legion_of_Congregated_Aspynarians,The_Republic_of_Beastistan,Terra_Invicta,Apachia," \
                        u"The_UCR,The_Chromeria,The_Empire_of_Joshpraise,Warbond,Rozi,The_Land_of_MS,Gode,Wunub," \
                        u"Thelema,James_Franco,The_Dominion_of_Northern_Ireland,Jumpy,United_North," \
                        u"The_Empire_Of_Haven,United_Sovereign_Nations,The_Pepe_Region,Owanet_Group,Zeus," \
                        u"The_French_Coast_Of_The_Somalis,Hamish_Rejects,Aurean_Sphere,The_Iron_League," \
                        u"The_Order_Of_The_Epidemic_X,Compass,Unity,Spear_Danes,United_Central,Baile,Walala_Dominion," \
                        u"Launching_Point,Praeceptor,Kurabul,Heiwa,The_National_Sovereignty_Coalition,Eggplant_City," \
                        u"A_Hole_To_Hide_In,Castlemaine_Military,Doll_Guldur,Communist,United_National_Federations," \
                        u"Amanuensis,Hand,Phantoms_of_Kaimera,The_Communist_Terrorists,Socialist_Pact," \
                        u"The_Holy_Empire_of_the_Papa_John,The_Lunar_Alliance,the_renewed_united_nations," \
                        u"Region_Of_Geography_Nerds,Peuce_Island,Carajo,The_Final_Fronteir," \
                        u"The_International_Congress_of_Nations,The_French_Empire,The_Greenup_Alliance," \
                        u"The_Nations_for_Freedom_and_Justice,Virtus,Myopia,Patria_Reborn,B1_Kingdom,Imaginary," \
                        u"The_Golden_Globe,Steptoe_Scorpions,Domon_Ord,The_Pareven_Isles,Space_Piracy_is_Legal," \
                        u"From_The_End,Waves,Country_of_God"

        # query template that will grab change events prior to cutoff point in late updating regions
        request_stem = u"https://www.nationstates.net/cgi-bin/api.cgi?q=happenings;filter=change;beforetime={" \
                       u"};view=region.{}"

        offset = 1

        while True:
            # determine yesterday, strip out time values. we'll add that in shortly.
            time_now = datetime.datetime.utcnow().replace(tzinfo=UTC(), hour=0, minute=0, second=0, microsecond=0) - \
                       datetime.timedelta(days=offset)

            # determine yesterday minor start and cutoff point
            start_minor = time_now + datetime.timedelta(hours=16)
            cutoff_minor = time_now + datetime.timedelta(hours=17)

            # determine yesterday major cutoff point
            start_major = time_now + datetime.timedelta(hours=4)
            cutoff_major = time_now + datetime.timedelta(hours=5)

            # store times here
            minor_time, major_time = 0, 0

            # determine minor update

            # generate request
            minor_request = request_stem.format(int(self.timestamp(cutoff_minor)), late_updaters)
            minor_query = urllib2.Request(minor_request, headers=headers)
            minor_xml = et.fromstring(urllib2.urlopen(minor_query).read())
            # print minor_request
            # iterate through queried events
            for event in minor_xml.iter(u'EVENT'):
                time = int(event.find(u"TIMESTAMP").text)
                kind = event.find(u"TEXT").text

                # if it's an influence update, we want it
                if u"influence" or u"ranked" in kind:
                    # event time - start time = update length
                    minor_time = time - int(self.timestamp(start_minor))
                    # print u"Event Text: ", kind
                    # print u"Event Time: ", datetime.datetime.fromtimestamp(time)
                    # print u"Calculated Minor Length: ", minor_time
                    break

            major_request = request_stem.format(int(self.timestamp(cutoff_major)), late_updaters)
            major_query = urllib2.Request(major_request, headers=headers)
            major_xml = et.fromstring(urllib2.urlopen(major_query).read())

            for event in major_xml.iter(u'EVENT'):
                time = int(event.find(u"TIMESTAMP").text)
                kind = event.find(u"TEXT").text
                if u"influence" in kind or u"ranked" in kind:
                    major_time = time - int(self.timestamp(start_major))
                    # print u"Event Text: ", kind
                    # print u"Event Time: ", datetime.datetime.fromtimestamp(time)
                    # print u"Calculated Major Length: ", major_time
                    break

            if minor_time < 0 or major_time < 0:
                offset += 1
            else:
                break

        return [minor_time, major_time]
