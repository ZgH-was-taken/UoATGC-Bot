import discord
from discord.ext import commands
from discord.utils import get
import random as rnd
import openpyxl
import asyncio
import datetime as dt


with open("token.txt", 'r') as f:
    token = f.readline()

with open("ids.txt", 'r') as f:
    roleMessageID = int(f.readline()[:-1])
    roleMessageResetID = int(f.readline())

client = discord.Client()
bot = commands.Bot(command_prefix='!', case_insensitive = True)
bot.remove_command('help')

wb = openpyxl.load_workbook('Member List.xlsx')
ws = wb['Paid Members 2019']


@bot.event
async def on_ready():
    global server, generalChannel, ruleChannel, botChannel, botMember, execBotChannel, execRole
    server = get(bot.guilds, name='UoATGC')
    generalChannel = get(server.channels, name='general')
    ruleChannel = get(server.channels, name='rules')
    botChannel = get(server.channels, name='bot-commands')
    botMember = get(server.members, id=bot.user.id)
    execBotChannel = get(server.channels, name='exec-bot')
    execRole = get(server.roles, name='Exec')
    print('Bot is ready')
    msg = await ruleChannel.fetch_message(roleMessageID)
    for reaction in msg.reactions:
        if reaction.emoji == '🎲':
            role = get(server.roles, name='Board Games')
            for user in await reaction.users().flatten():
                if user is not botMember: await user.add_roles(role)
        if reaction.emoji == '⚔️':
            role = get(server.roles, name='RPGs')
            for user in await reaction.users().flatten():
                if user is not botMember: await user.add_roles(role)
        if reaction.emoji == '🃏':
            role = get(server.roles, name='TCGs')
            for user in await reaction.users().flatten():
                if user is not botMember: await user.add_roles(role)
    msg = await ruleChannel.fetch_message(roleMessageResetID)
    for reaction in msg.reactions:
        if reaction.emoji == '🎲':
            role = get(server.roles, name='Board Games')
            for user in await reaction.users().flatten():
                if user is not botMember: await user.add_roles(role)
        if reaction.emoji == '⚔️':
            role = get(server.roles, name='RPGs')
            for user in await reaction.users().flatten():
                if user is not botMember: await user.add_roles(role)
        if reaction.emoji == '🃏':
            role = get(server.roles, name='TCGs')
            for user in await reaction.users().flatten():
                if user is not botMember: await user.add_roles(role)


welcomeMsg = '''Welcome to the TGC Discord! Be sure to read the server rules
Reply with your Uni ID to be given the member role, or react with 🎲
if you aren't from UoA to alert one of the exec members to message you
'''
@bot.event
async def on_member_join(user):
    await user.send(welcomeMsg)
    msg = await execBotChannel.send(user.name + ' has joined the server')
    await msg.add_reaction('🎲')


@bot.event
async def on_message(msg):
    if msg.guild is None:
        member = get(server.members, id=msg.author.id)
        if msg.author is bot.user or member is None: return
        elif get(member.roles, name='Member') is not None: return
        i = 2
        try:
            while ws.cell(i, 1).value is not None:
                if int(msg.content.strip()) == ws.cell(i, 2).value:
                    memberRole = get(server.roles, name='Member')
                    member = get(server.members, id=msg.author.id)
                    await member.add_roles(memberRole)
                    await generalChannel.send('Welcome, ' + member.mention + '!')
                    await execBotChannel.send(member.display_name + '/' + ws.cell(i,1).value + ' has joined as a member')
                    return
                i += 1
        except ValueError:
            return


@bot.event
async def on_raw_reaction_add(payload):
    guild = get(bot.guilds, id=payload.guild_id)
    emoji = payload.emoji.name
    if payload.user_id == botMember.id: return
    if guild is None:
        member = get(server.members, id=payload.user_id)
        if member is None: return
        if get(member.roles, name='Member') is None and emoji == '🎲':
            await execBotChannel.send(execRole.mention + member.display_name + ' wants to be manually made a member of the server')
        return
    channel = get(server.channels, id=payload.channel_id)
    member = payload.member
    bgRole = get(server.roles, name='Board Games')
    rpgRole = get(server.roles, name='RPGs')
    tcgRole = get(server.roles, name='TCGs')

    if get(member.roles, name='Member') is None: return
    if payload.message_id == roleMessageID:
        if emoji == '🎲':
            await member.add_roles(bgRole)
        elif emoji == '⚔️':
            await member.add_roles(rpgRole)
        elif emoji == '🃏':
            await member.add_roles(tcgRole)
        msg = await ruleChannel.fetch_message(roleMessageResetID)
        users = None
        for reaction in msg.reactions:
            if reaction.emoji == emoji:
                users = await reaction.users().flatten()
                break
        if users is not None:
            if member in users:
                await reaction.remove(member)
    if payload.message_id == roleMessageResetID:
        if emoji == '🎲':
            await member.add_roles(bgRole)
        elif emoji == '⚔️':
            await member.add_roles(rpgRole)
        elif emoji == '🃏':
            await member.add_roles(tcgRole)
        msg = await ruleChannel.fetch_message(roleMessageID)
        users = None
        for reaction in msg.reactions:
            if reaction.emoji == emoji:
                users = await reaction.users().flatten()
                break
        if users is not None:
            if member in users:
                await reaction.remove(member)
    if emoji == '🍆' or emoji == '🍑':
        msg = await channel.fetch_message(payload.message_id)
        if msg.author is botMember:
            await member.send('😉')
        


@bot.event
async def on_raw_reaction_remove(payload):
    guild = get(bot.guilds, id=payload.guild_id)
    if payload.user_id == botMember.id or guild is None: return
    emoji = payload.emoji.name
    channel = get(server.channels, id=payload.channel_id)
    member = get(server.members, id=payload.user_id)
    bgRole = get(server.roles, name='Board Games')
    rpgRole = get(server.roles, name='RPGs')
    tcgRole = get(server.roles, name='TCGs')

    if payload.message_id == roleMessageID:
        msg = await ruleChannel.fetch_message(roleMessageResetID)
        for reaction in msg.reactions:
            if reaction.emoji == emoji:
                users = await reaction.users().flatten()
                break
        if member in users: return
        if emoji == '🎲':
            await member.remove_roles(bgRole)
        elif emoji == '⚔️':
            await member.remove_roles(rpgRole)
        elif emoji == '🃏':
            await member.remove_roles(tcgRole)
    if payload.message_id == roleMessageResetID:
        msg = await ruleChannel.fetch_message(roleMessageID)
        for reaction in msg.reactions:
            if reaction.emoji == emoji:
                users = await reaction.users().flatten()
                break
        if member in users: return
        if emoji == '🎲':
            await member.remove_roles(bgRole)
        elif emoji == '⚔️':
            await member.remove_roles(rpgRole)
        elif emoji == '🃏':
            await member.remove_roles(tcgRole)


async def reset():
    while True:
        now = dt.datetime.now()
        midnight = dt.datetime.combine(dt.datetime.today() + dt.timedelta(days=1), dt.time.min)
        await asyncio.sleep((midnight-now).total_seconds() + 60)

        if rnd.random() < 0.01:
            await botMember.edit(nick='Carlos Webster, UoATGC President')
        else:
            await botMember.edit(nick=None)

        if dt.datetime.today().weekday() != 6: continue
        msg = await ruleChannel.fetch_message(roleMessageResetID)
        bgRole = get(server.roles, name='Board Games')
        rpgRole = get(server.roles, name='RPGs')
        tcgRole = get(server.roles, name='TCGs')
        for reaction in msg.reactions:
            emoji = reaction.emoji 
            for member in await reaction.users().flatten():
                if member.id == botMember.id: continue
                if emoji == '🎲':
                    await member.remove_roles(bgRole)
                elif emoji == '⚔️':
                    await member.remove_roles(rpgRole)
                elif emoji == '🃏':
                    await member.remove_roles(tcgRole)
                await reaction.remove(member)


bot.loop.create_task(reset())
bot.run(token)