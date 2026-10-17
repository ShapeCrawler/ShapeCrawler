using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.MediaContent;

internal sealed class AudioTiming(P.Picture picture)
{
    private const string PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";

    private P.Slide Slide => picture.Ancestors<P.Slide>().FirstOrDefault()
        ?? throw new SCException("Audio start mode is only supported on slides.");

    private string ShapeId => picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Id!.Value
        .ToString(CultureInfo.InvariantCulture);

    public AudioStartMode StartMode
    {
        get
        {
            this.EnsureAudioShape();
            var command = this.PlayCommand();
            if (command is null)
            {
                return AudioStartMode.WhenClickedOn;
            }

            if (command.Ancestors<P.CommonTimeNode>().Any(node => node.NodeType?.Value == P.TimeNodeValues.InteractiveSequence))
            {
                return AudioStartMode.WhenClickedOn;
            }

            var effect = command.Ancestors<P.CommonTimeNode>().First(node => node.PresetClass?.Value == P.TimeNodePresetClassValues.MediaCall);
            return effect.NodeType?.Value == P.TimeNodeValues.ClickEffect
                ? AudioStartMode.InClickSequence
                : AudioStartMode.Automatically;
        }
        set
        {
            if (!Enum.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }

            this.EnsureAudioShape();
            var startDelay = this.StartDelay.TotalMilliseconds.ToString("0", CultureInfo.InvariantCulture);
            var root = this.Root();
            var command = this.PlayCommand();
            if (command is not null)
            {
                // Remove only this media effect, leaving unrelated slide animations intact.
                var effect = command.Ancestors<P.CommonTimeNode>()
                    .First(node => node.PresetClass?.Value == P.TimeNodePresetClassValues.MediaCall);
                OpenXmlElement group = effect.Parent!;
                while (group.Parent is P.ChildTimeNodeList list && list.ChildElements.Count == 1 &&
                       list.Parent is P.CommonTimeNode parent && parent.NodeType is null)
                {
                    group = parent.Parent!;
                }

                var container = (P.ChildTimeNodeList)group.Parent!;
                group.Remove();
                if (!container.HasChildren)
                {
                    var parent = (P.CommonTimeNode)container.Parent!;
                    if (parent.NodeType?.Value == P.TimeNodeValues.InteractiveSequence)
                    {
                        parent.Parent!.Remove();
                    }
                    else
                    {
                        container.Remove();
                    }
                }
            }

            this.EnsureAudio(root);
            var main = value == AudioStartMode.WhenClickedOn ? null : this.MainSequence(root);
            var nextId = this.NextId();
            var effectType = value == AudioStartMode.InClickSequence ? "clickEffect" : "withEffect";
            var delay = value == AudioStartMode.InClickSequence ? "indefinite" : "0";
            var groupXml = $"""
                <p:par xmlns:p="{PresentationNamespace}"><p:cTn id="{nextId++}" fill="hold">
                  <p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>
                  <p:childTnLst><p:par><p:cTn id="{nextId++}" fill="hold">
                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                    <p:childTnLst><p:par><p:cTn id="{nextId++}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="{effectType}">
                      <p:stCondLst><p:cond delay="{startDelay}"/></p:stCondLst>
                      <p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr>
                        <p:cTn id="{nextId++}" dur="indefinite" fill="hold"/>
                        <p:tgtEl><p:spTgt spid="{this.ShapeId}"/></p:tgtEl>
                      </p:cBhvr></p:cmd></p:childTnLst>
                    </p:cTn></p:par></p:childTnLst>
                  </p:cTn></p:par></p:childTnLst>
                </p:cTn></p:par>
                """;
            var groupNode = new P.ParallelTimeNode(groupXml);
            if (command is not null)
            {
                // Keep the existing playback duration and command options.
                var preservedCommand = (P.Command)command.CloneNode(true);
                foreach (var node in preservedCommand.Descendants<P.CommonTimeNode>())
                {
                    node.Id = nextId++;
                }

                var generatedCommand = groupNode.Descendants<P.Command>().Single();
                generatedCommand.Parent!.ReplaceChild(preservedCommand, generatedCommand);
            }

            if (value == AudioStartMode.WhenClickedOn)
            {
                var sequence = new P.SequenceTimeNode($"""
                    <p:seq xmlns:p="{PresentationNamespace}" concurrent="1" nextAc="seek">
                      <p:cTn id="{nextId}" restart="whenNotActive" fill="hold" nodeType="interactiveSeq">
                        <p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="{this.ShapeId}"/></p:tgtEl></p:cond></p:stCondLst>
                        <p:childTnLst/>
                      </p:cTn>
                    </p:seq>
                    """);
                sequence.CommonTimeNode!.ChildTimeNodeList!.Append(groupNode);
                root.ChildTimeNodeList!.Append(sequence);
            }
            else
            {
                main!.ChildTimeNodeList!.Append(groupNode);
            }
        }
    }

    public TimeSpan StartDelay
    {
        get
        {
            this.EnsureAudioShape();
            var delay = this.PlayEffect()?.StartConditionList?.GetFirstChild<P.Condition>()?.Delay?.Value;
            return uint.TryParse(delay, NumberStyles.None, CultureInfo.InvariantCulture, out var milliseconds)
                ? TimeSpan.FromMilliseconds(milliseconds)
                : TimeSpan.Zero;
        }
        set
        {
            if (value < TimeSpan.Zero || value.TotalMilliseconds > uint.MaxValue)
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }

            this.EnsureAudioShape();
            if (this.PlayCommand() is null)
            {
                this.StartMode = AudioStartMode.WhenClickedOn;
            }

            var effect = this.PlayEffect()!;
            effect.StartConditionList ??= new P.StartConditionList();
            var condition = effect.StartConditionList.GetFirstChild<P.Condition>()
                ?? effect.StartConditionList.AppendChild(new P.Condition());
            condition.Delay = Math.Round(value.TotalMilliseconds).ToString("0", CultureInfo.InvariantCulture);
        }
    }

    private P.CommonTimeNode? PlayEffect() => this.PlayCommand()?.Ancestors<P.CommonTimeNode>()
        .First(node => node.PresetClass?.Value == P.TimeNodePresetClassValues.MediaCall);

    private void EnsureAudioShape()
    {
        if (!picture.Descendants<DocumentFormat.OpenXml.Drawing.AudioFromFile>().Any())
        {
            throw new SCException("Start mode is only supported for audio content.");
        }
    }

    private P.Command? PlayCommand() => this.Slide.Timing?.Descendants<P.Command>()
        .FirstOrDefault(command => command.CommandName?.Value?.StartsWith("play", StringComparison.Ordinal) == true &&
            command.Descendants<P.ShapeTarget>().Any(target => target.ShapeId?.Value == this.ShapeId));

    private uint NextId() => (this.Slide.Timing?.Descendants<P.CommonTimeNode>()
        .Select(node => node.Id?.Value ?? 0U).DefaultIfEmpty().Max() ?? 0U) + 1;

    private P.CommonTimeNode Root()
    {
        var timing = this.Slide.Timing;
        if (timing is null)
        {
            timing = new P.Timing();
            this.Slide.AddChild(timing, true);
        }

        timing.TimeNodeList ??= new P.TimeNodeList();
        var root = timing.Descendants<P.CommonTimeNode>().FirstOrDefault(node => node.NodeType?.Value == P.TimeNodeValues.TmingRoot);
        if (root is null)
        {
            root = new P.CommonTimeNode
            {
                Id = this.NextId(), Duration = "indefinite", Restart = P.TimeNodeRestartValues.Never,
                NodeType = P.TimeNodeValues.TmingRoot, ChildTimeNodeList = new P.ChildTimeNodeList()
            };
            timing.TimeNodeList.Append(new P.ParallelTimeNode(root));
        }

        root.ChildTimeNodeList ??= new P.ChildTimeNodeList();
        return root;
    }

    private P.CommonTimeNode MainSequence(P.CommonTimeNode root)
    {
        var main = root.Descendants<P.CommonTimeNode>().FirstOrDefault(node => node.NodeType?.Value == P.TimeNodeValues.MainSequence);
        if (main is null)
        {
            var sequence = new P.SequenceTimeNode($"""
                <p:seq xmlns:p="{PresentationNamespace}" concurrent="1" nextAc="seek">
                  <p:cTn id="{this.NextId()}" dur="indefinite" nodeType="mainSeq"><p:childTnLst/></p:cTn>
                  <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
                  <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
                </p:seq>
                """);
            root.ChildTimeNodeList!.PrependChild(sequence);
            main = sequence.CommonTimeNode!;
        }

        main.ChildTimeNodeList ??= new P.ChildTimeNodeList();
        return main;
    }

    private void EnsureAudio(P.CommonTimeNode root)
    {
        if (this.Slide.Timing!.Descendants<P.Audio>().Any(audio =>
                audio.Descendants<P.ShapeTarget>().Any(target => target.ShapeId?.Value == this.ShapeId)))
        {
            return;
        }

        root.ChildTimeNodeList!.Append(new P.Audio($"""
            <p:audio xmlns:p="{PresentationNamespace}"><p:cMediaNode vol="80000">
              <p:cTn id="{this.NextId()}" fill="hold" display="0">
                <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
                <p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>
              </p:cTn>
              <p:tgtEl><p:spTgt spid="{this.ShapeId}"/></p:tgtEl>
            </p:cMediaNode></p:audio>
            """));
    }
}

