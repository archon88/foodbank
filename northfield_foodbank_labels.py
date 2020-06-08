#!/usr/bin/env python
# coding: utf-8

import glob, os, sys

from PIL import Image, ImageDraw, ImageFont, ImageFile
from IPython.display import display
import numpy as np
import pandas as pd

ImageFile.LOAD_TRUNCATED_IMAGES = True

pd.set_option('display.max_colwidth', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)


class RecipientData:
    '''
    Class to contain the data for a single addressee and ensure
    appropriate serialisation of the data to string.
    '''
    def __init__(self, row):
        self.name = row['name']
        self.phone = row['phone number']
        self.address = row['address']
        self.postcode = row['postcode']
        self.age = row['age']
        self.adults = row['number of adults']
        self.elderly = row['how many of these adults are over 70']
        self.children = row['number of children']
        self.teens = row['how many of these children are over 12']
        self.vegetarian = row['vegetarian, halal or kosher']
        self.allergy = row['specify allergy if they have one']
        self.cooker = row['do they have a cooker']
        self.hob = row['do they have a hob']
        self.kettle = row['do they have a kettle']
        self.microwave = row['do they have a microwave']
        self.notes = row['notes']
        
    def __str__(self):
        label = f"{self.name}\n{self.address}\n{self.postcode}"
        label += f"\nPhone: {self.phone}" # are we always showing the number now?
        label += f"\nAge: {self.age}"
        label += f"\nAdults: {self.adults} (Elderly: {self.elderly})"
        label += f"\nChildren: {self.children} (Teens: {self.teens})"
        label += f"\nNo cooker" if not self.cooker else ''
        label += f"\nNo hob" if not self.hob else ''
        label += f"\nNo kettle" if not self.kettle else ''
        label += f"\nNo microwave" if not self.microwave else ''
        label += f"\nNotes:" if (self.notes and str(self.notes).casefold() != 'nan') else ''
        return label


class DataSet:
    '''
    A simple wrapper around a dataframe to contain custom information.
    '''
    def __init__(self, datafile, dropna=True):
        self.df = pd.read_excel(datafile)
        self.df = self.drop_empty_rows(self.df) if dropna else self.df
        self.df = self.fix_colnames(self.df)
        self.labels = [RecipientData(row) for (index, row) in self.df.iterrows()]
    
    @staticmethod
    def fix_colnames(df):
        '''
        Standardise column names and typing.
        '''
        cols = {colname: colname.casefold().replace('?', '').replace('  ', ' ') for colname in df.columns}
        df = df.rename(columns=cols)
        # This is not the most elegant solution but workable for now
        df['phone number'] = df['phone number'].astype(str).apply(lambda x: x.replace(' ', '')).apply(lambda x: f'0{x.replace(".0", "")}') #if not x.startswith('0') else x.replace(".0", ""))
        for col in ['age', 'number of adults', 'how many of these adults are over 70', 'number of children', 'how many of these children are over 12']:
            df[col] = df[col].astype(pd.Int64Dtype())
        for col in ['vegetarian, halal or kosher', 'do they have a cooker', 'do they have a hob', 'do they have a kettle', 'do they have a microwave']:  # boolean conversion
                df[col] = df[col].apply(lambda x: True if x.casefold() == 'yes' else False)
        return df
    
    @staticmethod
    def drop_empty_rows(df):
        return df.dropna(subset=['Name'])


def write_label(labeldir, text, index, vegetarian, allergy):
    """
    Create a single "A7" PNG from a row of spreadsheet data.
    """
    (W, H) = 297, 210  # A-series paper aspect ratio
    img = Image.new('RGB', (W, H), color = 'white')
    d = ImageDraw.Draw(img)
    w, h = d.textsize(text)
    font_normal = ImageFont.truetype("/usr/share/fonts/truetype/artemisia/GFSArtemisia.otf", size=14)
    # font_large = ImageFont.truetype("/usr/share/fonts/truetype/artemisia/GFSArtemisia.otf", size=20)
    font_huge = ImageFont.truetype("/usr/share/fonts/truetype/artemisia/GFSArtemisia.otf", size=22)
    d.text(((W-w) // 10, (H-h) // 2), text=text, font=font_normal, fill='black')  # left-align
    if vegetarian:
        bottom = H - 25
        d.text((10, bottom), text='V', font=font_huge, fill='green')
    if allergy is not None and str(allergy) != 'nan':
        bottom = H - 25
        d.text(((W-w) // 2, bottom), text=allergy.upper(), font=font_huge, fill='red')
    img.save(f'{labeldir}/test_label_{index}.png')


def concatenate_image_page(savedir, labs, fname):
    """
    Combine labels into 2x4 grids which will be "A4" pages.
    """
    new_im = Image.new('RGB', (2*297, 4*210))  # the canvas
    blank_label = Image.new('RGB', (297, 210), color = 'white')  # to cover empty spaces
    ctr = 0
    for i in range(0, 2*297, 297):
        for j in range(0, 4*210, 210):
            try:
                new_im.paste(labs[ctr], (i,j))
                ctr += 1
            except IndexError: # if number of images is not 8
                new_im.paste(blank_label, (i,j))
                ctr += 1
    return new_im


def make_all_pages(labdir, savedir):
    """
    Top-level function, to run over all labels and generate a single PDF of printable A4 sheets.
    """
    assert labdir != savedir, f'Warning, cannot write to directory containing label files!'
    all_labs = [Image.open(f'{labdir}/{l}') for l in os.listdir(labdir) if ".png" in l and "label" in l]
    chunked = [all_labs[i:i + 8] for i in range(0, len(all_labs), 8)]
    pagectr = 1
    pages = []
    for chunk in chunked:
        p = concatenate_image_page(savedir, chunk, f'concatenated_labels_page_{pagectr}')
        pages.append(p)
        pagectr += 1
    pages[0].save(f"{savedir}/out.pdf", save_all=True, append_images=pages[1:])


if __name__ == '__main__':
    BASEDIR = input('Enter top-level directory:\n')
    spreadsheets = glob.glob(f'{BASEDIR}/*xlsx')
    assert len(spreadsheets) == 1, 'Error, number of spreadsheets in directory is not one!'
    d = DataSet(f'{BASEDIR}/Label Spreadsheet 8-5-20.xlsx')
    os.makedirs(f'{BASEDIR}/labels/a4labs', exist_ok=True)
    os.makedirs(f'{BASEDIR}/labels/output', exist_ok=True)
    for i, lab in enumerate(d.labels):
        write_label(f'{BASEDIR}/labels/a4labs/', str(lab), i, vegetarian=lab.vegetarian, allergy=lab.allergy)
    make_all_pages(f'{BASEDIR}/labels/a4labs/', f'{BASEDIR}/labels/output')