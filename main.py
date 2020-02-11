import discord
from discord.ext import commands
from discord.utils import get
import random as rnd
import openpyxl
import asyncio


with open("token.txt", 'r') as f:
    token = f.readline()[:-1]
    botID = f.readline()[:-1]
    serverName = f.readline()

client = discord.Client()
bot = commands.Bot(command_prefix='!', case_insensitive = True)
bot.remove_command('help')

wb = openpyxl.load_workbook('Member List.xlsx')
ws = wb['Paid Members 2019']


@bot.event
async def on_ready():
    global server, generalChannel, ruleChannel, botChannel, execBotChannel, execRole
    server = get(bot.guilds, name=serverName)
    generalChannel = get(server.channels, name='general')
    ruleChannel = get(server.channels, name='rules')
    botChannel = get(server.channels, name='bot-commands')
    execBotChannel = get(server.channels, name='exec-bot')
    execRole = get(server.roles, name='Exec')
    print('Bot is ready')
    msg = await ruleChannel.send('Alternatively, react to this message to gain the role until the end of the next session')
    await msg.add_reaction(emoji='🎲')
    await msg.add_reaction(emoji='⚔️')
    await msg.add_reaction(emoji='🃏')


welcomeMsg = '''Welcome to the TGC Discord!
[Welcome message]
Reply with your Uni ID to be given the member role:
'''
@bot.event
async def on_member_join(user):
    await user.send(welcomeMsg)
    await execBotChannel.send(user.name + ' has joined the server')


@bot.event
async def on_message(msg):
    if msg.guild is None:
        member = get(server.members, id=msg.author.id)
        if msg.author is bot.user or member is None:
            return
        elif get(member.roles, name='Exec') is not None:
            return
        i = 2
        try:
            while ws.cell(i, 2).value is not None:
                if int(msg.content.strip()) == ws.cell(i, 2).value:
                    execRole = get(server.roles, name='Exec')
                    member = get(server.members, id=msg.author.id)
                    await member.add_roles(execRole)
                    await generalChannel.send('Welcome, ' + member.mention + '!')
                    await execBotChannel.send(member.display_name + '/' + ws.cell(i,1).value + ' has joined as a member')
                    return
                i += 1
        except ValueError:
            return


@bot.event
async def on_reaction_add(reaction, user):
    if reaction.message.guild is None:
        member = get(server.members, id=user.id)
        if member is None or user is bot.user:
            return
        if get(member.roles, name='Member') is None and reaction.emoji == '🎲':
            await execBotChannel.send('@ exec ' + member.display_name + ' wants to be manually made a member of the server')
            await execBotChannel.send(execRole.mention + member.display_name + ' wants to be manually made a member of the server')



bot.run(token)